local M = {}

local function health_fn(name)
    local health = vim.health or {}
    if type(health[name]) == 'function' then
        return health[name]
    end

    local legacy_name = 'report_' .. name
    if type(health[legacy_name]) == 'function' then
        return health[legacy_name]
    end

    return function(...) end
end

local health_start = health_fn('start')
local health_ok = health_fn('ok')
local health_warn = health_fn('warn')
local health_error = health_fn('error')

local function report_executable(name, required, help_msg)
    if vim.fn.executable(name) == 1 then
        health_ok(string.format('`%s` found', name))
        return true
    end

    if required then
        health_error(string.format('`%s` is required but not found in PATH', name), help_msg)
    else
        health_warn(string.format('`%s` not found (optional)', name), help_msg)
    end
    return false
end

function M.check()
    health_start('docxedit')

    health_start('Core dependencies')
    report_executable('zip', true, 'Install `zip` and ensure it is in PATH.')
    report_executable('unzip', true, 'Install `unzip` and ensure it is in PATH.')

    health_start('Formatting')
    report_executable(
        'xmlformat',
        false,
        'Without `xmlformat`, XML remains unformatted but plugin still works.'
    )

    health_start('Word reload integration')
    local uname = vim.loop.os_uname()
    local sysname = uname.sysname
    if sysname == 'Windows_NT' then
        report_executable(
            'powershell',
            true,
            'Install PowerShell and ensure it is available in PATH.'
        )
    elseif sysname == 'Darwin' then
        report_executable(
            'osascript',
            true,
            'macOS should provide `osascript`; ensure command-line tools are available.'
        )
        health_ok('Grant Automation permissions for your terminal/Neovim when prompted.')
    else
        local msg = string.format('OS `%s` is not officially supported for Word reload.', sysname)
        local advice = 'Live Word reload currently targets Windows and macOS.'
        health_warn(msg, advice)
    end
end

return M
