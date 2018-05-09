# xlsx2ssv Introduction

`xlsx2csv`をラップしてxlsxファイルをスペース区切りファイル(SSV)に変換する

# Requirement

* [xlsx2csv](https://github.com/dilshod/xlsx2csv)
* bash
* gnu grep(macだとBSD Grepがデフォルトなのでそのままだと使えません)

# Install

`$ cp -p xlsx2ssv /usr/local/bin/`

# Usage

`$ xlsx2ssv sheetNo cellrange-start cellrange-end xlsxfile > stdout`
`$ xlsx2ssv sheetNo cellrange-start cellrange-end < xlsxfile > stdout`

# Example

`$ xlsx2ssv 1 A1 AC13 file.xlsx > output`

# Test

batsを利用。テストにはOpen-uspコマンドを使用している

* [Open-usp-Tukubai](https://github.com/usp-engineers-community/Open-usp-Tukubai)
