#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
STDOUT.sync = true

require 'win32ole'
require 'time'

# Excel VBA定数のロード
module Excel; end

def init_excel()
  # Excelオブジェクト生成
  excel = WIN32OLE.new('Excel.Application')
  excel.visible = false
  # 上書きメッセージを抑制
  excel.displayAlerts = false

  WIN32OLE.const_load(excel, Excel)

  return excel
end

=begin
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
=end

def read_excel(excel, file, sheet_num = 1)
  p 'ちよつきんのたいきん時間を入力してね！'
  go_home_time = gets
  p 'きようのしゆつきんしかんをにゆうりよくしてね！'
  attendance_time = gets
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)
  today = Time.now()

  # 列ごとに処理
  #sheet.UsedRange.Rows.each do |row|
    # セルごとに処理

    sheet.range('A10:A40').each do |cell|

      t = cell.value

      if today.day == t.day then
        c = cell.Address.to_s
        
        puts sheet.range(c.gsub(/A/, 'C')).value = time
        puts cell.Address
        puts today.day - 1
      end
      
    #row.Columns.each do |cell|
      #end
    end
      
    # 保存
    book.saveAs(file)

    # ファイルを閉じる
    book.close

  end  

def main()
  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  #file = fso.GetAbsolutePathName('./sample.xlsx')
  file = fso.GetAbsolutePathName('C:/Users/HMP01156/OUT/ロンテック勤務表(2019年02月)(福留).xlsm')

  excel = init_excel()

  #create_excel(excel, file)
  read_excel(excel, file)

  excel.quit()
end

main()
