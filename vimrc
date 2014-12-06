if has("autocmd")
  au BufReadPost * if line("'\"") > 1 && line("'\"") <= line("$") | exe "normal! g'\"" | endif
endif
colorscheme evening
set et
set nu
set ai
set incsearch
set backup
syntax on
