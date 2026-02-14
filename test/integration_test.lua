vim.opt.rtp:append(vim.fn.getcwd())

local docxedit = require('docxedit')

local tests = {}

local function test(name, fn)
    table.insert(tests, { name = name, fn = fn })
end

local function assert_true(condition, message)
    if condition then return end
    error(message or 'assertion failed')
end

local function with_stub(target, key, replacement, fn)
    local original = target[key]
    target[key] = replacement
    local ok, result = pcall(fn)
    target[key] = original
    if not ok then error(result) end
    return result
end

local function escape_for_windows(path)
    path = path:gsub('/', '\\')
    path = path:gsub('`', '``')
    path = path:gsub('"', '`"')
    path = path:gsub('%$', '`$')
    path = path:gsub('\r', ' ')
    path = path:gsub('\n', ' ')
    return path
end

local function escape_for_applescript(path)
    path = path:gsub('\\', '\\\\')
    path = path:gsub('"', '\\"')
    path = path:gsub('\r', '\\r')
    path = path:gsub('\n', '\\n')
    return path
end

test('powershell script escapes metacharacters in doc path', function()
    local command
    local raw_path = '/tmp/p$ath`"name.docx'
    local normalized = vim.fn.fnamemodify(raw_path, ':p')
    local expected = escape_for_windows(normalized)

    with_stub(vim.loop, 'os_uname', function()
        return { sysname = 'Windows_NT', version = 'Windows' }
    end, function()
        with_stub(vim, 'system', function(cmd)
            command = cmd
            return {
                wait = function() return { code = 0, stdout = '', stderr = '' } end,
            }
        end, function()
            docxedit.reload_docx('zipfile://' .. raw_path .. '::word/document.xml')
        end)
    end)

    assert_true(type(command) == 'table', 'expected command argv')
    local script_path = command[#command]
    assert_true(vim.fn.filereadable(script_path) == 1, 'expected generated script')
    local script_body = table.concat(vim.fn.readfile(script_path), '\n')
    local expected_line = '$docPath = "' .. expected .. '"'
    assert_true(script_body:find(expected_line, 1, true), 'expected escaped powershell path')
end)

test('applescript escapes metacharacters in doc path', function()
    local command
    local raw_path = '/tmp/a\\b"c\nfile.docx'
    local normalized = vim.fn.fnamemodify(raw_path, ':p')
    local expected = escape_for_applescript(normalized)

    with_stub(vim.loop, 'os_uname', function()
        return { sysname = 'Darwin', version = 'Darwin Kernel Version' }
    end, function()
        with_stub(vim, 'system', function(cmd)
            command = cmd
            return {
                wait = function() return { code = 0, stdout = '', stderr = '' } end,
            }
        end, function()
            docxedit.reload_docx('zipfile://' .. raw_path .. '::word/document.xml')
        end)
    end)

    assert_true(type(command) == 'table', 'expected command argv')
    local script_path = command[#command]
    assert_true(vim.fn.filereadable(script_path) == 1, 'expected generated script')
    local script_body = table.concat(vim.fn.readfile(script_path), '\n')
    local expected_line = 'set unixPath to "' .. expected .. '"'
    assert_true(script_body:find(expected_line, 1, true), 'expected escaped applescript path')
end)

test('stopping watcher cleans generated script file', function()
    local command
    local base_path = vim.fn.tempname() .. '.docx'
    local bufname = 'zipfile://' .. base_path .. '::word/document.xml'
    vim.fn.writefile({ 'docx' }, base_path)

    local bufnr = vim.api.nvim_create_buf(false, true)
    vim.api.nvim_buf_set_name(bufnr, bufname)

    with_stub(vim, 'system', function(cmd)
        command = cmd
        return {
            wait = function() return { code = 0, stdout = '', stderr = '' } end,
        }
    end, function()
        with_stub(vim.loop, 'new_fs_event', function()
            local watcher = {}
            function watcher:start() return true end
            function watcher:stop() end
            function watcher:close() end
            return watcher
        end, function()
            docxedit.reload_docx(bufname)
            local script_path = command[#command]
            assert_true(vim.fn.filereadable(script_path) == 1, 'expected script before stop')

            docxedit.watch_edit(bufname)
            assert_true(docxedit.stop_watch(bufname), 'expected watcher stop')
            assert_true(vim.fn.filereadable(script_path) == 0, 'expected script cleanup')
        end)
    end)

    vim.api.nvim_buf_delete(bufnr, { force = true })
end)

local failures = 0
for _, item in ipairs(tests) do
    local ok, err = pcall(item.fn)
    if ok then
        print('ok - ' .. item.name)
    else
        failures = failures + 1
        io.stderr:write('not ok - ' .. item.name .. '\n')
        io.stderr:write(err .. '\n')
    end
end

if failures > 0 then
    vim.cmd('cquit ' .. failures)
end

print(string.format('all integration tests passed (%d)', #tests))
