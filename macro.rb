#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
STDOUT.sync = true

require 'win32ole'
require 'time'
require 'mail'

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

def greeting(today)

  now = today.strftime('%Y年 %m月 %x日 (%a)')
  judge_hour = today.hour

  if judge_hour >= 8 and judge_hour <= 12　then
    greeting = 'おはよう！\n今日は ' + now
  elsif judge_hour >=13 and judge_hour <= 16 then
    greeting = 'こんにちは！\n今日は ' + now
  elsif judge_hour == 17 then
    greeting = 'お疲れ様！\n今日は ' + now
  end

  puts greeting
end

def send_mail(excel, file)

  Mail.defaults do
    delivery_method :smtp, {
      :address => 'smtp.gmail.com',
      :port => 587,
      :domain => 'example.com',
      :user_name => "#{mail_from}",
      :password => "#{mail_passwd}",
      :authentication => :login,
      :enable_starttls_auto => true
    }
  end

  mail = Mail.new do
    from    'from@example.co.jp'
    to      'to@example.co.jp'
    subject 'subject text'
    body    ''
    add_file './sample/xlxs'
  end
  
end


def update_excel(excel, file, sheet_num = 1, today)
  
  puts '直近の退勤時間を入力してね！'
  tmp_end_time = gets
  
  puts '今日の出勤時間を入力してね！'
  tmp_start_time = gets
  
  go_home_time = Time.parse(tmp_end_time)
  attendance_time = Time.parse(tmp_start_time)
  
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)
  
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

  today = Time.now()

  greeting(today)
  
  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  #file = fso.GetAbsolutePathName('./sample.xlsx')
  file = fso.GetAbsolutePathName('sample.xlsm')

  excel = init_excel()

  send_mail(excel, file)
  update_excel(excel, file, today)

  excel.quit()
end

main()

puts 'END'
sleep(3)