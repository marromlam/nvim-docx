local M = {}

--- Checks if the operating system is Windows.
--- @return boolean True if the OS is Windows, false otherwise.
local function is_windows()
    -- Check if the OS is Windows
    return vim.loop.os_uname().version:match('Windows')
end

--- Escapes the given file path based on the operating system.
-- On Windows, it replaces forward slashes with backslashes.
-- On other systems, it returns the path unchanged.
-- @param path string: The file path to escape.
-- @return string: The escaped file path.
local function escape_path(path)
    if is_windows() then
        return path:gsub('/', '\\')
    else
        return path
    end
end

--- Generates the appropriate script to open a DOCX file in Microsoft Word based on the operating system.
-- @param docx_path string: The path to the DOCX file.
-- @return string: The command to execute the script.
local function get_os_script(docx_path)
    local script_path = vim.fn.stdpath('config') .. '/docxedit_tmp_script'
    local full_script = ''

    if is_windows() then
        full_script = string.format(
            [[
            $docPath = "%s"
            $word = New-Object -ComObject Word.Application
            $word.Visible = $true
            foreach ($doc in $word.Documents) {
              if ($doc.FullName -eq $docPath) {
                $doc.Close([ref]$false)
                break
              }
            }
            $word.Documents.Open($docPath)
        ]],
            escape_path(docx_path)
        )
        vim.fn.writefile(vim.split(full_script, '\n'), script_path .. '.ps1')
        return 'powershell -ExecutionPolicy Bypass -File '
            .. script_path
            .. '.ps1'
    else
        full_script = string.format(
            [[
            set unixPath to "%s"
            set macPath to POSIX file unixPath as alias
            tell application "Microsoft Word"
              set theDoc to open macPath
              activate macPath
            end tell
            tell application "Microsoft Word"
              close active document saving no
              open macPath
            end tell
        ]],
            docx_path
        )
        vim.fn.writefile(vim.split(full_script, '\n'), script_path .. '.scpt')
        return 'osascript ' .. script_path .. '.scpt'
    end
end

--- Edits a DOCX file by extracting its contents and setting up an auto command to update the document.
-- @param docx_path string: The path to the DOCX file.
M.edit_docx = function(docx_path)
    if not docx_path or docx_path == '' then
        print('Usage: :EditDocx path/to/file.docx')
        return
    end

    local docx_abs = vim.fn.fnamemodify(docx_path, ':p')
    local base = vim.fn.fnamemodify(docx_path, ':t:r')
    local temp_dir = vim.fn.stdpath('cache') .. '/docxedit/' .. base

    vim.fn.mkdir(temp_dir, 'p')
    os.execute(string.format('unzip -o "%s" -d "%s"', docx_abs, temp_dir))

    vim.api.nvim_create_autocmd('BufWritePost', {
        pattern = temp_dir .. '/word/document.xml',
        callback = function()
            local zip_cmd =
                string.format('cd "%s" && zip -r "%s" *', temp_dir, docx_abs)
            os.execute(zip_cmd)
            os.execute(get_os_script(docx_abs))
            print('Updated Word document.')
        end,
        once = false,
    })

    vim.cmd('edit ' .. temp_dir .. '/word/document.xml')
end

--- Reloads a DOCX file in the buffer.
-- @param bufname string: The name of the buffer to reload.
M.reload_docx = function(bufname)
    bufname = bufname or vim.api.nvim_buf_get_name(0)
    local base_path, xml_path = bufname:match('^zipfile://(.-)::(.*)$')
    os.execute(get_os_script(base_path))
end

--- Watches a file for changes and reloads it in the buffer.
--- @param bufname string: The name of the buffer to reload.
--- @param base_path string: The path to the file or directory to watch.
local function file_watcher(bufname, base_path)
    local watcher = vim.loop.new_fs_event()
    if not watcher then return end
    watcher:start(base_path, {}, function(err)
        if err then return end

        vim.schedule(function()
            vim.cmd('edit!' .. bufname)
            vim.cmd(':silent %!xmlformat')
            if watcher then watcher:stop() end
            file_watcher(bufname, base_path)
        end)
    end)
end

--- Sets up a watcher to edit a DOCX file and reload it upon changes.
-- @param bufname string: The name of the buffer to watch.
M.watch_edit = function(bufname)
    bufname = bufname or vim.api.nvim_buf_get_name(0)
    local base_path, xml_path = bufname:match('^zipfile://(.-)::(.*)$')
    if not base_path then return end
    file_watcher(bufname, base_path)
end

return M
