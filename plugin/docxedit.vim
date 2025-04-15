lua << EOF
vim.api.nvim_create_user_command("EditDocx", function(opts)
  require("docxedit").edit_docx(opts.args)
end, { nargs = 1 })
EOF
