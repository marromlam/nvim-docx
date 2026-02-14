# nvim-docx

A small Neovim plugin for editing `.docx` XML and reloading the file in Microsoft Word.

https://github.com/user-attachments/assets/2b777285-b506-4a4d-97bf-8f30627c196f

## What it does

- Works with `zipfile://...::word/document.xml` buffers
- Watches for external updates and reloads the buffer
- Runs `xmlformat` if installed (otherwise leaves XML as-is)
- Reopens the document in Word on macOS/Windows

## Install (lazy.nvim)

```lua
{
  "marromlam/nvim-docx",
  lazy = true,
}
```

## Configuration

All options are optional:

```lua
require("docxedit").setup({
  reload_debounce_ms = 150,     -- debounce for filesystem events
  auto_format = true,           -- run xmlformat when available
  debug = false,                -- emit debug logs with vim.notify
  cleanup_scripts_on_stop = true, -- remove temp reload scripts on stop/exit
  cleanup_stale_scripts_on_load = true, -- delete old temp scripts on startup/setup
  stale_script_ttl_seconds = 86400, -- max age (seconds) before temp scripts are deleted
})
```

## Usage

Open the document XML directly:

```vim
:edit zipfile:///absolute/path/to/file.docx::word/document.xml
```

## Keybinding Configuration

### Option 1: `lazy.nvim` `keys` config

```lua
{
  "marromlam/nvim-docx",
  keys = {
    {
      "<leader>X",
      function()
        require("docxedit").start_watch()
      end,
      mode = { "n", "i" },
      desc = "Reload MS Word",
    },
    {
      "<leader>x",
      function()
        require("docxedit").stop_watch()
      end,
      mode = "n",
      desc = "Stop docx watch",
    },
  },
}
```

### Option 2: plain `init.lua` keymaps

```lua
vim.keymap.set({ 'n', 'i' }, '<leader>X', function()
  require('docxedit').start_watch()
end, { desc = 'Reload MS Word' })
vim.keymap.set('n', '<leader>x', function()
  require('docxedit').stop_watch()
end, { desc = 'Stop docx watch' })
```

## Requirements

- `zip` and `unzip` in your system PATH
- `xmlformat` is optional, but recommended since Word XML is usually one long line
- Microsoft Word installed
- PowerShell (Windows) or AppleScript (macOS)

If you want `xmlformat` on macOS, this formula works well:
[homebrew-custom/xmlformat.rb](https://github.com/marromlam/homebrew-custom/blob/main/xmlformat.rb)

## Healthcheck

Run:

```vim
:checkhealth docxedit
```

This validates core tools (`zip`, `unzip`), optional formatter (`xmlformat`), and OS-specific reload support (`powershell`/`osascript`).

## Vim Help

After install, open:

```vim
:help docxedit
```

When developing locally from this repo, generate tags once:

```vim
:helptags /absolute/path/to/nvim-docx/doc
```
