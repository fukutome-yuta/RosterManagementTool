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
  
  puts '直近の退勤時間を入力してね！'
  tmp_end_time = gets
  
  puts '今日の出勤時間を入力してね！'
  tmp_start_time = gets
  
  go_home_time = Time.parse(tmp_end_time)
  attendance_time = Time.parse(tmp_start_time)
  
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)
  
  today = Time.now()
  last_day = today.day - 1

    sheet.range('A10:A40').each do |cell|

      tmp_day = cell.value

      if last_day == tmp_day.day then
        
        tmp_end_cell = cell.Address.to_s
        go_home_cell = sheet.range(tmp_end_cell.gsub(/A/, 'D'))
        target_cell_No = tmp_end_cell.delete("^0-9").to_i

        while go_home_cell.value == nil do
          
          target_cell_No = target_cell_No - 1
          target_cell_address = '$D$' + target_cell_No.to_s
          go_home_cell = sheet.range(target_cell_address)

        end

        go_home_cell.value = go_home_time

      end

      if today.day == tmp_day.day then

        tmp_start_cell = cell.Address.to_s
        attendance_cell = sheet.range(tmp_start_cell.gsub(/A/, 'C'))
        attendance_cell.value = attendance_time

      end
      
    end
        
    # 保存
    book.saveAs(file)

    # ファイルを閉じる
    book.close
        
    puts '更新完了！'

  end  

def main()
  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  #file = fso.GetAbsolutePathName('./sample.xlsx')
  file = fso.GetAbsolutePathName('sample.xlsm')

  excel = init_excel()

  #create_excel(excel, file)
  read_excel(excel, file)

  excel.quit()
end

main()

puts 'END'
sleep(3)
