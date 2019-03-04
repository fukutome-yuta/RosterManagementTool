require 'win32ole'

app = WIN32OLE.new('Excel.Application')
book = app.Workbooks.Open(app.GetOpenFilename)

#使っているワークシート範囲を一行ずつ取り出す
for row in book.ActiveSheet.UsedRange.Rows do
  #取り出した行から、セルを一つづつ取り出す
  for cell in row.Columns do
    p cell.Address
    p cell.Value
    p '-------'
  end
end

book.close(false)
app.quit

#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
STDOUT.sync = true

require 'win32ole'

# Excel VBA定数のロード
module Excel; end

def init_excel()
  # Excelオブジェクト生成
  excel = WIN32OLE.new('Excel.Application')
  excel.visible = true
  # 上書きメッセージを抑制
  excel.displayAlerts = false

  WIN32OLE.const_load(excel, Excel)

  return excel
end

def create_excel(excel, file)
  # 新規ブックを作成
  workbook = excel.workbooks.add

  # 先頭シートを選択
  sheet = workbook.sheets[1]

  # 九九の表を作成
  (1..9).each do |i|
    sheet.rows[1].columns[i + 1] = i
    sheet.rows[i + 1].columns[1] = i
  end
  sheet.range('B2:J10').value = '=$A2*B$1'

  # ボーダーライン
  sheet.range('A1:J10').borders.lineStyle = Excel::XlContinuous

  # 表のヘッダー
  range = sheet.range('A1:A10,B1:J1')
  # 背景色
  range.interior.themeColor = Excel::XlThemeColorAccent1
  # フォント
  range.font.themeColor = Excel::XlThemeColorDark1
  range.font.bold = true

  # 列の幅
  sheet.columns('A:J').columnWidth = 6

  # 保存
  workbook.saveAs(file)

  # ファイルを閉じる
  workbook.close
end

def read_excel(excel, file, sheet_num = 1)
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)

  # 列ごとに処理
  sheet.UsedRange.Rows.each do |row|
    # セルごとに処理
    row.Columns.each do |cell|
      puts cell.value
    end
  end
end

def main()
  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  file = fso.GetAbsolutePathName('./sample.xlsx')

  excel = init_excel()

  create_excel(excel, file)
  read_excel(excel, file)

  excel.quit()
end

main()
