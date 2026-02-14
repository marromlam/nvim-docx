local M = {}
local watchers_by_buf = {}
local script_files_by_path = {}

local config = {
    reload_debounce_ms = 150,
    auto_format = true,
    debug = false,
    cleanup_scripts_on_stop = true,
    cleanup_stale_scripts_on_load = true,
    stale_script_ttl_seconds = 86400,
}

local function is_windows()
    return vim.loop.os_uname().version:match('Windows')
end

local function escape_path(path)
    if is_windows() then
        return path:gsub('/', '\\')
    else
        return path
    end
end

local function log_debug(message)
    if not config.debug then return end
    vim.notify('[docxedit] ' .. message, vim.log.levels.DEBUG)
end

local function log_warn(message)
    vim.notify('[docxedit] ' .. message, vim.log.levels.WARN)
end

local function escape_powershell_string(str)
    if not str then return '' end
    str = str:gsub('`', '``')
    str = str:gsub('"', '`"')
    str = str:gsub('%$', '`$')
    str = str:gsub('\r', ' ')
    str = str:gsub('\n', ' ')
    return str
end

local function escape_applescript_string(str)
    if not str then return '' end
    str = str:gsub('\\', '\\\\')
    str = str:gsub('"', '\\"')
    str = str:gsub('\r', '\\r')
    str = str:gsub('\n', '\\n')
    return str
end

local function write_file_safely(path, lines, context)
    local ok, write_err = pcall(vim.fn.writefile, lines, path)
    if ok then return true end
    vim.notify(
        string.format('Failed writing %s: %s', context, write_err),
        vim.log.levels.ERROR
    )
    return false
end

local function validate_docx_path(docx_path)
    if not docx_path or docx_path == '' then
        return nil, 'invalid document path: empty'
    end

    if docx_path:find('..', 1, true) then
        return nil, 'invalid document path: path traversal detected'
    end

    if docx_path:find('\0', 1, true) then
        return nil, 'invalid document path: contains NUL byte'
    end

    local is_absolute = docx_path:match('^/') or docx_path:match('^%a:[/\\]')
    if not is_absolute then
        return nil, 'invalid document path: must be absolute'
    end

    return vim.fn.fnamemodify(docx_path, ':p')
end

local function run_command(command, context)
    local output = ''
    local exit_code = 0

    if vim.system then
        local result = vim.system(command, { text = true }):wait()
        exit_code = result.code or 0
        output = result.stderr or result.stdout or ''
    else
        output = vim.fn.system(command)
        exit_code = vim.v.shell_error
    end

    if exit_code ~= 0 then
        vim.notify(
            string.format(
                '%s failed (exit %d): %s',
                context,
                exit_code,
                vim.trim(output)
            ),
            vim.log.levels.ERROR
        )
        return false
    end

    return true
end

local function cleanup_script_file(docx_path)
    local normalized_path = vim.fn.fnamemodify(docx_path or '', ':p')
    local script_path = script_files_by_path[normalized_path]
    if not script_path then return end
    script_files_by_path[normalized_path] = nil
    if not config.cleanup_scripts_on_stop then return end
    if vim.fn.filereadable(script_path) == 1 then
        vim.fn.delete(script_path)
    end
end

local function cleanup_stale_scripts(max_age_seconds)
    local script_dir = vim.fn.stdpath('cache') .. '/docxedit'
    if vim.fn.isdirectory(script_dir) ~= 1 then return end

    local files = vim.fn.glob(script_dir .. '/docxedit_tmp_script_*', false, true)
    local now = os.time()
    for _, file in ipairs(files) do
        local mtime = vim.fn.getftime(file)
        if mtime > 0 and (now - mtime) > max_age_seconds then
            vim.fn.delete(file)
            log_debug('Removed stale script: ' .. file)
        end
    end
end

local function stop_watcher(bufnr)
    local watcher_state = watchers_by_buf[bufnr]
    if not watcher_state then return end

    watcher_state.cancelled = true

    if watcher_state.watcher then
        pcall(watcher_state.watcher.stop, watcher_state.watcher)
        pcall(watcher_state.watcher.close, watcher_state.watcher)
    end

    if watcher_state.cleanup_autocmd_id then
        pcall(vim.api.nvim_del_autocmd, watcher_state.cleanup_autocmd_id)
    end

    if watcher_state.debounce_timer then
        pcall(watcher_state.debounce_timer.stop, watcher_state.debounce_timer)
        pcall(watcher_state.debounce_timer.close, watcher_state.debounce_timer)
    end

    if watcher_state.base_path then
        log_debug('Stopping watcher for ' .. watcher_state.base_path)
        cleanup_script_file(watcher_state.base_path)
    end

    watchers_by_buf[bufnr] = nil
end

local function doc_hash(docx_path)
    local ok, hash = pcall(vim.fn.sha256, docx_path)
    if ok and type(hash) == 'string' and hash ~= '' then
        return hash:sub(1, 12)
    end

    return tostring(vim.loop.hrtime())
end

local function get_os_script(docx_path)
    local normalized_path, path_err = validate_docx_path(docx_path)
    if not normalized_path then
        vim.notify(path_err, vim.log.levels.ERROR)
        return nil
    end

    local script_dir = vim.fn.stdpath('cache') .. '/docxedit'
    local script_ext = is_windows() and '.ps1' or '.scpt'
    local script_path =
        string.format('%s/docxedit_tmp_script_%s%s', script_dir, doc_hash(normalized_path), script_ext)
    local full_script = ''
    vim.fn.mkdir(script_dir, 'p')
    script_files_by_path[normalized_path] = script_path

    if vim.fn.filereadable(script_path) == 1 then
        log_debug('Reusing reload script: ' .. script_path)
        if is_windows() then
            return {
                'powershell',
                '-ExecutionPolicy',
                'Bypass',
                '-File',
                script_path,
            }
        end
        return { 'osascript', script_path }
    end

    if is_windows() then
        local safe_path = escape_powershell_string(escape_path(normalized_path))
        full_script = string.format(
            [[
            $docPath = "%s"
            $word = $null
            try {
              $word = New-Object -ComObject Word.Application
              $word.Visible = $true

              foreach ($doc in $word.Documents) {
                if ($doc.FullName -eq $docPath) {
                  $doc.Close([ref]$false) | Out-Null
                  break
                }
              }

              $word.Documents.Open($docPath) | Out-Null
            } catch {
              Write-Error "Failed to open document: $_"
              exit 1
            }
        ]],
            safe_path
        )
        if not write_file_safely(script_path, vim.split(full_script, '\n'), 'PowerShell reload script') then
            return nil
        end
        return {
            'powershell',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            script_path,
        }
    else
        local safe_path = escape_applescript_string(normalized_path)
        full_script = string.format(
            [[
            set unixPath to "%s"
            set macPath to POSIX file unixPath as alias
            tell application "System Events"
              set frontAppName to name of first process whose frontmost is true
            end tell
            tell application "Microsoft Word"
              try
                close active document saving no
              end try
              open macPath
            end tell
            tell application frontAppName to activate
        ]],
            safe_path
        )
        if not write_file_safely(script_path, vim.split(full_script, '\n'), 'AppleScript reload script') then
            return nil
        end
        return { 'osascript', script_path }
    end
end

M.reload_docx = function(bufname)
    bufname = bufname or vim.api.nvim_buf_get_name(0)
    local base_path = bufname:match('^zipfile://(.-)::.*$')
    if not base_path then
        vim.notify(
            'reload_docx expects a zipfile buffer path.',
            vim.log.levels.WARN
        )
        return
    end

    local command = get_os_script(base_path)
    if not command then return end
    log_debug('Reloading DOCX in Word: ' .. base_path)
    run_command(command, 'Word reload')
end

local function refresh_buffer(bufnr)
    local state = watchers_by_buf[bufnr]
    if state and state.cancelled then return end

    if not vim.api.nvim_buf_is_valid(bufnr) then
        stop_watcher(bufnr)
        return
    end

    local winid = vim.fn.bufwinid(bufnr)
    if winid == -1 then return end

    vim.api.nvim_win_call(winid, function()
        local previous_view = vim.fn.winsaveview()
        local current_bufname = vim.api.nvim_buf_get_name(bufnr)
        local escaped_bufname = vim.fn.fnameescape(current_bufname)
        local ok_edit, edit_err = pcall(vim.cmd, 'edit! ' .. escaped_bufname)
        if not ok_edit then
            vim.notify(
                string.format('Failed to reload DOCX buffer: %s', edit_err),
                vim.log.levels.ERROR
            )
            return
        end

        if config.auto_format and vim.fn.executable('xmlformat') == 1 then
            vim.cmd('silent %!xmlformat')
            if vim.v.shell_error ~= 0 then
                vim.notify(
                    string.format('xmlformat failed (exit %d).', vim.v.shell_error),
                    vim.log.levels.ERROR
                )
            end
        elseif config.auto_format then
            log_debug('xmlformat not found; skipping formatting')
        end

        local last_line = vim.api.nvim_buf_line_count(bufnr)
        previous_view.lnum = math.min(previous_view.lnum, last_line)
        vim.fn.winrestview(previous_view)
    end)
end

local function file_watcher(bufname, base_path)
    local bufnr = vim.fn.bufnr(bufname)
    if bufnr < 0 then return end

    local current = watchers_by_buf[bufnr]
    if current and current.base_path == base_path then return end

    stop_watcher(bufnr)
    log_debug('Starting watcher for ' .. base_path)

    local watcher = vim.loop.new_fs_event()
    if not watcher then
        vim.notify('Failed to create file watcher.', vim.log.levels.ERROR)
        return
    end

    local debounce_timer = vim.loop.new_timer()
    if not debounce_timer then
        watcher:close()
        vim.notify('Failed to create debounce timer.', vim.log.levels.ERROR)
        return
    end

    watchers_by_buf[bufnr] = {
        watcher = watcher,
        base_path = base_path,
        debounce_timer = debounce_timer,
        cancelled = false,
    }

    local cleanup_autocmd_id = vim.api.nvim_create_autocmd(
        { 'BufDelete', 'BufWipeout' },
        {
            buffer = bufnr,
            once = true,
            callback = function() stop_watcher(bufnr) end,
        }
    )
    watchers_by_buf[bufnr].cleanup_autocmd_id = cleanup_autocmd_id

    local ok_start, start_err = pcall(
        watcher.start,
        watcher,
        base_path,
        {},
        function(err)
            if err then return end

            local state = watchers_by_buf[bufnr]
            if not state or state.cancelled or not state.debounce_timer then
                return
            end

            state.debounce_timer:stop()
            state.debounce_timer:start(config.reload_debounce_ms, 0, function()
                vim.schedule(function()
                    local inner_state = watchers_by_buf[bufnr]
                    if not inner_state or inner_state.cancelled then return end
                    refresh_buffer(bufnr)
                end)
            end)
        end
    )
    if not ok_start then
        stop_watcher(bufnr)
        vim.notify(
            string.format('Failed to start file watcher: %s', start_err),
            vim.log.levels.ERROR
        )
        return
    end
end

M.watch_edit = function(bufname)
    bufname = bufname or vim.api.nvim_buf_get_name(0)
    local base_path = bufname:match('^zipfile://(.-)::.*$')
    if not base_path then return end
    file_watcher(bufname, base_path)
end

M.start_watch = function(bufname)
    bufname = bufname or vim.api.nvim_buf_get_name(0)
    local ok_write, write_err = pcall(vim.cmd, 'write')
    if not ok_write then
        vim.notify(
            string.format('Failed to write buffer before start_watch: %s', write_err),
            vim.log.levels.ERROR
        )
        return false
    end

    M.reload_docx(bufname)
    M.watch_edit(bufname)
    return true
end

M.setup = function(opts)
    opts = opts or {}

    local normalized = vim.deepcopy(opts)

    if normalized.reload_debounce_ms ~= nil then
        if type(normalized.reload_debounce_ms) ~= 'number' or normalized.reload_debounce_ms < 0 then
            log_warn('Invalid `reload_debounce_ms`; keeping previous value.')
            normalized.reload_debounce_ms = nil
        else
            normalized.reload_debounce_ms = math.floor(normalized.reload_debounce_ms)
        end
    end

    if normalized.stale_script_ttl_seconds ~= nil then
        if type(normalized.stale_script_ttl_seconds) ~= 'number' or normalized.stale_script_ttl_seconds < 0 then
            log_warn('Invalid `stale_script_ttl_seconds`; keeping previous value.')
            normalized.stale_script_ttl_seconds = nil
        else
            normalized.stale_script_ttl_seconds = math.floor(normalized.stale_script_ttl_seconds)
        end
    end

    local bool_fields = {
        'auto_format',
        'debug',
        'cleanup_scripts_on_stop',
        'cleanup_stale_scripts_on_load',
    }
    for _, key in ipairs(bool_fields) do
        if normalized[key] ~= nil and type(normalized[key]) ~= 'boolean' then
            log_warn(string.format('Invalid `%s`; expected boolean. Keeping previous value.', key))
            normalized[key] = nil
        end
    end

    config = vim.tbl_extend('force', config, normalized)

    if config.cleanup_stale_scripts_on_load then
        cleanup_stale_scripts(config.stale_script_ttl_seconds)
    end
end

M.stop_watch = function(bufname)
    local bufnr
    if type(bufname) == 'number' then
        bufnr = bufname
    elseif type(bufname) == 'string' and bufname ~= '' then
        bufnr = vim.fn.bufnr(bufname)
    else
        bufnr = vim.api.nvim_get_current_buf()
    end

    if bufnr < 0 or not watchers_by_buf[bufnr] then return false end
    stop_watcher(bufnr)
    return true
end

M.cleanup = function()
    local bufnrs = {}
    for bufnr, _ in pairs(watchers_by_buf) do
        table.insert(bufnrs, bufnr)
    end

    for _, bufnr in ipairs(bufnrs) do
        stop_watcher(bufnr)
    end

    if config.cleanup_scripts_on_stop then
        for path, _ in pairs(script_files_by_path) do
            cleanup_script_file(path)
        end
    end
end

vim.api.nvim_create_autocmd('VimLeavePre', {
    callback = function() M.cleanup() end,
})

if config.cleanup_stale_scripts_on_load then
    cleanup_stale_scripts(config.stale_script_ttl_seconds)
end

return M
