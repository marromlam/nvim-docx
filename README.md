# nvim-docx

Edit `.docx` files directly from Neovim by modifying their XML, and have changes live-reloaded into Microsoft Word.




https://github.com/user-attachments/assets/2b777285-b506-4a4d-97bf-8f30627c196f





## ✨ Features

- Unzips `.docx` and opens the XML content (`word/document.xml`)
- Automatically re-zips on save
- Reloads the file in Microsoft Word on Windows or macOS

## 🚀 Usage

```vim
:EditDocx path/to/your.docx
```

## 📦 Installation (lazy.nvim)

```lua
{
  "yourusername/nvim-docx-edit",
  config = function()
    -- nothing to configure yet
  end
}
```

## 🖥️ Requirements

- `zip` and `unzip` in your system PATH
- Microsoft Word installed
- PowerShell (Windows) or AppleScript (macOS)
