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

test('reload_docx warns on non-zipfile buffers', function()
    local notifications = {}
    with_stub(vim, 'notify', function(message, level)
        table.insert(notifications, { message = message, level = level })
    end, function()
        docxedit.reload_docx('/tmp/not-a-zipfile.xml')
    end)

    assert_true(#notifications == 1, 'expected one warning notification')
    assert_true(
        notifications[1].level == vim.log.levels.WARN,
        'expected WARN level notification'
    )
    assert_true(
        notifications[1].message:match('zipfile buffer'),
        'expected zipfile warning message'
    )
end)

test('reload_docx blocks path traversal input', function()
    local notifications = {}
    local command_calls = 0

    with_stub(vim, 'notify', function(message, level)
        table.insert(notifications, { message = message, level = level })
    end, function()
        with_stub(vim, 'system', function(command, _)
            command_calls = command_calls + 1
            return {
                wait = function()
                    return { code = 0, stdout = '', stderr = '' }
                end,
            }
        end, function()
            docxedit.reload_docx('zipfile://../../../etc/passwd::word/document.xml')
        end)
    end)

    assert_true(command_calls == 0, 'expected no command execution for traversal path')
    assert_true(#notifications >= 1, 'expected traversal path notification')
    assert_true(
        notifications[1].message:match('path traversal detected'),
        'expected path traversal error'
    )
end)

test('start_watch writes and triggers reload/watch', function()
    local calls = {
        write = 0,
        reload = 0,
        watch = 0,
    }
    local target_buf = 'zipfile:///tmp/example.docx::word/document.xml'

    with_stub(vim, 'cmd', function(command)
        if command == 'write' then
            calls.write = calls.write + 1
            return
        end
    end, function()
        with_stub(docxedit, 'reload_docx', function(bufname)
            calls.reload = calls.reload + 1
            assert_true(bufname == target_buf, 'expected start_watch buffer passthrough')
        end, function()
            with_stub(docxedit, 'watch_edit', function(bufname)
                calls.watch = calls.watch + 1
                assert_true(bufname == target_buf, 'expected watch_edit buffer passthrough')
            end, function()
                local ok = docxedit.start_watch(target_buf)
                assert_true(ok, 'expected start_watch success')
            end)
        end)
    end)

    assert_true(calls.write == 1, 'expected one write')
    assert_true(calls.reload == 1, 'expected one reload')
    assert_true(calls.watch == 1, 'expected one watch')
end)

test('start_watch returns false when write fails', function()
    local notified = false
    local called_reload = false
    local called_watch = false

    with_stub(vim, 'cmd', function(command)
        if command == 'write' then error('simulated write failure') end
    end, function()
        with_stub(vim, 'notify', function()
            notified = true
        end, function()
            with_stub(docxedit, 'reload_docx', function()
                called_reload = true
            end, function()
                with_stub(docxedit, 'watch_edit', function()
                    called_watch = true
                end, function()
                    local ok = docxedit.start_watch(
                        'zipfile:///tmp/example.docx::word/document.xml'
                    )
                    assert_true(not ok, 'expected start_watch failure')
                end)
            end)
        end)
    end)

    assert_true(notified, 'expected error notification')
    assert_true(not called_reload, 'expected no reload on write failure')
    assert_true(not called_watch, 'expected no watch on write failure')
end)

test('reload_docx uses hashed script path per document', function()
    local command_calls = {}
    local temp_root = vim.fn.tempname()
    local path_one = temp_root .. '/a/same-name.docx'
    local path_two = temp_root .. '/b/same-name.docx'
    vim.fn.mkdir(temp_root .. '/a', 'p')
    vim.fn.mkdir(temp_root .. '/b', 'p')
    vim.fn.writefile({ 'one' }, path_one)
    vim.fn.writefile({ 'two' }, path_two)

    with_stub(vim, 'system', function(command, _)
        table.insert(command_calls, command)
        return {
            wait = function()
                return { code = 0, stdout = '', stderr = '' }
            end,
        }
    end, function()
        docxedit.reload_docx('zipfile://' .. path_one .. '::word/document.xml')
        docxedit.reload_docx('zipfile://' .. path_two .. '::word/document.xml')
    end)

    assert_true(#command_calls == 2, 'expected two reload command invocations')
    assert_true(
        type(command_calls[1]) == 'table',
        'expected argv table command execution'
    )

    local script_one = command_calls[1][#command_calls[1]]
    local script_two = command_calls[2][#command_calls[2]]
    local hash_one = vim.fn.sha256(path_one):sub(1, 12)
    local hash_two = vim.fn.sha256(path_two):sub(1, 12)

    assert_true(script_one:match(hash_one), 'expected hash for first document path')
    assert_true(script_two:match(hash_two), 'expected hash for second document path')
    assert_true(script_one ~= script_two, 'expected unique script paths')
    assert_true(vim.fn.filereadable(script_one) == 1, 'expected first script file')
    assert_true(vim.fn.filereadable(script_two) == 1, 'expected second script file')

    if command_calls[1][1] == 'osascript' then
        local lines = vim.fn.readfile(script_one)
        local script_body = table.concat(lines, '\n')
        assert_true(
            script_body:match('frontAppName'),
            'expected front app preservation in AppleScript'
        )
        assert_true(
            not script_body:match('activate macPath'),
            'expected AppleScript to avoid forcing Word focus'
        )
    end
end)

test('watch_edit is idempotent and cleans watcher on buffer wipeout', function()
    local start_calls = 0
    local stop_calls = 0
    local close_calls = 0

    local bufnr = vim.api.nvim_create_buf(false, true)
    local base_path = vim.fn.tempname() .. '.docx'
    local buffer_name = 'zipfile://' .. base_path .. '::word/document.xml'
    vim.fn.writefile({ 'docx placeholder' }, base_path)
    vim.api.nvim_buf_set_name(bufnr, buffer_name)

    with_stub(vim.loop, 'new_fs_event', function()
        local watcher = {}
        function watcher:start()
            start_calls = start_calls + 1
            return true
        end
        function watcher:stop()
            stop_calls = stop_calls + 1
        end
        function watcher:close()
            close_calls = close_calls + 1
        end

        return watcher
    end, function()
        docxedit.watch_edit(buffer_name)
        docxedit.watch_edit(buffer_name)
        assert_true(start_calls == 1, 'expected single watcher start for same buffer')

        vim.api.nvim_buf_delete(bufnr, { force = true })
    end)

    assert_true(stop_calls == 1, 'expected watcher stop on buffer wipeout')
    assert_true(close_calls == 1, 'expected watcher close on buffer wipeout')
end)

test('watch_edit ignores non-zipfile buffers', function()
    local created = 0
    with_stub(vim.loop, 'new_fs_event', function()
        created = created + 1
        return {}
    end, function()
        docxedit.watch_edit('/tmp/plain-file.xml')
    end)

    assert_true(created == 0, 'expected no watcher creation for plain buffers')
end)

test('health module is available for :checkhealth docxedit', function()
    local ok, health = pcall(require, 'docxedit.health')
    assert_true(ok, 'expected docxedit.health module')
    assert_true(type(health.check) == 'function', 'expected health.check function')
end)

test('watch_edit debounces repeated fs events', function()
    local on_change
    local edit_calls = 0
    local original_cmd = vim.cmd
    local original_path = vim.env.PATH
    local bufnr = vim.api.nvim_create_buf(false, true)
    local base_path = vim.fn.tempname() .. '.docx'
    local buffer_name = 'zipfile://' .. base_path .. '::word/document.xml'
    vim.fn.writefile({ 'docx placeholder' }, base_path)
    vim.api.nvim_buf_set_name(bufnr, buffer_name)
    vim.api.nvim_set_current_buf(bufnr)

    with_stub(vim.loop, 'new_fs_event', function()
        local watcher = {}
        function watcher:start(_, _, callback)
            on_change = callback
            return true
        end
        function watcher:stop() end
        function watcher:close() end
        return watcher
    end, function()
        with_stub(vim, 'cmd', function(command)
            if command:match('^edit! ') then
                edit_calls = edit_calls + 1
                return original_cmd(command)
            end
        end, function()
            vim.env.PATH = '/nonexistent'
            docxedit.watch_edit(buffer_name)
            assert_true(type(on_change) == 'function', 'expected watcher callback')

            on_change(nil)
            on_change(nil)
            on_change(nil)

            vim.wait(400, function() return edit_calls > 0 end, 20)
            assert_true(edit_calls == 1, 'expected a single debounced refresh')
        end)
    end)

    vim.env.PATH = original_path
    docxedit.stop_watch(buffer_name)
    vim.api.nvim_buf_delete(bufnr, { force = true })
end)

test('watch_edit skips xmlformat when unavailable', function()
    local on_change
    local xmlformat_called = false
    local original_cmd = vim.cmd
    local original_path = vim.env.PATH
    local bufnr = vim.api.nvim_create_buf(false, true)
    local base_path = vim.fn.tempname() .. '.docx'
    local buffer_name = 'zipfile://' .. base_path .. '::word/document.xml'
    vim.fn.writefile({ 'docx placeholder' }, base_path)
    vim.api.nvim_buf_set_name(bufnr, buffer_name)
    vim.api.nvim_set_current_buf(bufnr)

    with_stub(vim.loop, 'new_fs_event', function()
        local watcher = {}
        function watcher:start(_, _, callback)
            on_change = callback
            return true
        end
        function watcher:stop() end
        function watcher:close() end
        return watcher
    end, function()
        with_stub(vim, 'cmd', function(command)
            if command == 'silent %!xmlformat' then
                xmlformat_called = true
                return
            end
            return original_cmd(command)
        end, function()
            vim.env.PATH = '/nonexistent'
            docxedit.watch_edit(buffer_name)
            assert_true(type(on_change) == 'function', 'expected watcher callback')

            on_change(nil)
            vim.wait(400, function() return true end, 20)
        end)
    end)

    vim.env.PATH = original_path
    assert_true(not xmlformat_called, 'expected xmlformat to be skipped')
    docxedit.stop_watch(buffer_name)
    vim.api.nvim_buf_delete(bufnr, { force = true })
end)

test('setup can disable auto_format', function()
    local on_change
    local xmlformat_called = false
    local original_cmd = vim.cmd
    local bufnr = vim.api.nvim_create_buf(false, true)
    local base_path = vim.fn.tempname() .. '.docx'
    local buffer_name = 'zipfile://' .. base_path .. '::word/document.xml'
    vim.fn.writefile({ 'docx placeholder' }, base_path)
    vim.api.nvim_buf_set_name(bufnr, buffer_name)
    vim.api.nvim_set_current_buf(bufnr)

    docxedit.setup({ auto_format = false })
    with_stub(vim.loop, 'new_fs_event', function()
        local watcher = {}
        function watcher:start(_, _, callback)
            on_change = callback
            return true
        end
        function watcher:stop() end
        function watcher:close() end
        return watcher
    end, function()
        with_stub(vim, 'cmd', function(command)
            if command == 'silent %!xmlformat' then
                xmlformat_called = true
                return
            end
            return original_cmd(command)
        end, function()
            docxedit.watch_edit(buffer_name)
            assert_true(type(on_change) == 'function', 'expected watcher callback')
            on_change(nil)
            vim.wait(400, function() return true end, 20)
        end)
    end)

    docxedit.stop_watch(buffer_name)
    docxedit.setup({ auto_format = true })
    assert_true(not xmlformat_called, 'expected xmlformat disabled by setup')
    vim.api.nvim_buf_delete(bufnr, { force = true })
end)

test('setup warns and ignores invalid option types', function()
    local warnings = {}
    with_stub(vim, 'notify', function(message, level)
        if level == vim.log.levels.WARN then
            table.insert(warnings, message)
        end
    end, function()
        docxedit.setup({
            reload_debounce_ms = -1,
            stale_script_ttl_seconds = 'old',
            auto_format = 'yes',
            debug = 1,
            cleanup_scripts_on_stop = 'no',
            cleanup_stale_scripts_on_load = 'maybe',
        })
    end)

    assert_true(#warnings >= 5, 'expected warnings for invalid setup values')
    docxedit.setup({
        reload_debounce_ms = 150,
        stale_script_ttl_seconds = 86400,
        auto_format = true,
        debug = false,
        cleanup_scripts_on_stop = true,
        cleanup_stale_scripts_on_load = true,
    })
end)

test('setup removes stale scripts when enabled', function()
    local script_dir = vim.fn.stdpath('cache') .. '/docxedit'
    vim.fn.mkdir(script_dir, 'p')
    local stale_script = script_dir .. '/docxedit_tmp_script_stale.scpt'
    vim.fn.writefile({ 'old-script' }, stale_script)

    if vim.loop.fs_utime then
        local old_time = os.time() - 3600
        vim.loop.fs_utime(stale_script, old_time, old_time)
    end

    assert_true(vim.fn.filereadable(stale_script) == 1, 'expected stale script fixture')
    docxedit.setup({
        cleanup_stale_scripts_on_load = true,
        stale_script_ttl_seconds = 1,
    })
    assert_true(vim.fn.filereadable(stale_script) == 0, 'expected stale script cleanup')
end)

test('stop_watch manually stops active watchers', function()
    local stop_calls = 0
    local close_calls = 0
    local bufnr = vim.api.nvim_create_buf(false, true)
    local base_path = vim.fn.tempname() .. '.docx'
    local buffer_name = 'zipfile://' .. base_path .. '::word/document.xml'
    vim.fn.writefile({ 'docx placeholder' }, base_path)
    vim.api.nvim_buf_set_name(bufnr, buffer_name)

    with_stub(vim.loop, 'new_fs_event', function()
        local watcher = {}
        function watcher:start() return true end
        function watcher:stop()
            stop_calls = stop_calls + 1
        end
        function watcher:close()
            close_calls = close_calls + 1
        end
        return watcher
    end, function()
        docxedit.watch_edit(buffer_name)
        assert_true(docxedit.stop_watch(buffer_name), 'expected active watcher stop')
        assert_true(not docxedit.stop_watch(buffer_name), 'expected false when already stopped')
    end)

    assert_true(stop_calls == 1, 'expected one watcher stop call')
    assert_true(close_calls == 1, 'expected one watcher close call')
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

print(string.format('all tests passed (%d)', #tests))
