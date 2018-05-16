#!/usr/bin/env bats 

# TEST_Requirement [Open-usp-Tukubai](https://github.com/usp-engineers-community/Open-usp-Tukubai)

setup() {
  cd $BATS_TEST_DIRNAME
}

@test "xlsxファイルがない" {
    run xlsx2ssv 1 A1 C3 null.xlsx

    [ "$status" -eq 1 ]
    [ "$output" = "存在しないxlsxファイルです" ]
}

@test "1シート目のA1からC3まで取得(セル小文字表記)" {
    run  xlsx2ssv 1 a1 c3 valid.xlsx

    [ "$status" -eq 0 ]
    echo "$output" > output
    [ "$(gyo output)" = "3" ]
    [ "$(retu output)" = "3" ]
}

@test "1シート目のA1からC3まで取得(セル大文字表記)" {
     run xlsx2ssv 1 A1 C3 valid.xlsx 

    [ "$status" -eq 0 ]
    echo "$output" > output
    [ "$(gyo output)" = "3" ]
    [ "$(retu output)" = "3" ]
}

@test "2シート目のB2からD10まで取得" {
     run xlsx2ssv 2 b2 D10 valid.xlsx

    [ "$status" -eq 0 ]
    echo "$output" > output
    [ "$(gyo output)" = "9" ]
    [ "$(retu output)" = "3" ]
}

@test "1シート目のA1からC3まで取得(セル小文字表記: 標準入力版)" {
     run  xlsx2ssv 1 a1 c3 < valid.xlsx
  
     [ "$status" -eq 0 ]
     echo "$output" > output
     [ "$(gyo output)" = "3" ]
     [ "$(retu output)" = "3" ]
}


@test "2シート目のB2からD10まで取得(標準入力版)" {
      run xlsx2ssv 2 b2 D10 < valid.xlsx
  
     [ "$status" -eq 0 ]
     echo "$output" > output
     [ "$(gyo output)" = "9" ]
     [ "$(retu output)" = "3" ]
} 

@test "1シート目のA1からC3まで取得(パイプによる標準入力版)" {
     run bash -c 'cat valid.xlsx | xlsx2ssv 1 a1 c3'
  
     [ "$status" -eq 0 ]
     echo "$output" > output
     [ "$(gyo output)" = "3" ]
     [ "$(retu output)" = "3" ]
} 
  
  
@test "2シート目のB2からD10まで取得(パイプによる標準入力版)" {
      run bash -c 'cat valid.xlsx | xlsx2ssv 2 b2 D10'
  
     [ "$status" -eq 0 ]
     echo "$output" > output
     [ "$(gyo output)" = "9" ]
     [ "$(retu output)" = "3" ]
} 

@test "1シート目のD2からGの有効行まで取得" {
    skip
     run xlsx2ssv 2 D2 G- valid2.xlsx
    [ "$status" -eq 0 ]
    #未対応とします
}

@test "2文字アルファベットセルが正しいか1" {
    run xlsx2ssv 2 cc1 cc1 valid.xlsx

    [ "$status" -eq 0 ]
    [ "$output" = "col81" ]
}

@test "2文字アルファベットセルが正しいか2" {
    run xlsx2ssv 2 DE1 DE1 valid.xlsx

    [ "$status" -eq 0 ]
    [ "$output" = "col109" ]
}

@test "引数が足りない" {
    run xlsx2ssv a1 c3 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "${lines[0]}" = "必要な引数は4つです" ]
    [ "${lines[1]}" = "e.g. xlsx2ssv 1 A1 AC13 file.xlsx > output" ]

}

@test "引数が足りない(標準入力があるときは3つでもOK)" {
    run xlsx2ssv  1 a1 c3 < valid.xlsx 
    [ "$status" -eq 0 ]
}

@test "存在しないシートを選択" {
     run xlsx2ssv 5 a1 c3 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "${lines[0]}" = "xlsx2csvエラー" ]
    [ "${lines[1]}" = "Sheet 5 not found" ]
}

@test "xlsxではないファイルを選択" {
     run xlsx2ssv 1 a1 c3 invalid.xlsx 
    [ "$status" -eq 1 ]

    [ "${lines[0]}" = "xlsx2csvエラー" ]
    [ "${lines[1]}" = "Invalid xlsx file: invalid.xlsx" ]
}

@test "セル範囲が逆(逆には対応しない)" {
     run xlsx2ssv 1 a3 c1 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "セル範囲異常" ]

     run xlsx2ssv 1 c1 a3 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "セル範囲異常" ]

     run xlsx2ssv 1 D9 B8 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "セル範囲異常" ]

     run xlsx2ssv 2 AA31 Z32 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "セル範囲異常" ]

     run xlsx2ssv 2 Z32 AA31 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "セル範囲異常" ]
}

@test "セル範囲が不正(2回マッチ)" {
     run xlsx2ssv 1 a1 b3b3 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "アルファベットマッチ回数異常" ]

     run xlsx2ssv 1 a1 3b3 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "アルファベットマッチ回数異常" ]

     run xlsx2ssv 1 a1 b3b valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "アルファベットマッチ回数異常" ]
}

@test "セル範囲が不正(横はzzまでにしか対応しない)" {
     run xlsx2ssv 1 a1 aaa3 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "アルファベット文字数異常" ]
}

@test "セル範囲が不正(アルファベット欠落)" {
     run xlsx2ssv 1 a1 33 valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "アルファベットマッチ回数異常" ]
}

@test "セル範囲が不正(数字欠落)" {
     run xlsx2ssv 1 a1 z valid.xlsx 
    [ "$status" -eq 1 ]

    [ "$output" = "数字マッチ回数異常" ]
}

@test "後始末(テストではない)" {
    rm output
}

