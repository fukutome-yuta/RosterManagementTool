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

  now = today.strftime('%Y年 %m月 %d日 (%a)')
  judge_hour = today.hour

  if judge_hour >= 8 and judge_hour <= 12
    @greeting = "おはよう！\n今日は " + now
    @closing_remarks = '今日も一日頑張ろう！'
  elsif judge_hour >=13 and judge_hour <= 16
    @greeting = "こんにちは！\n今日は " + now
    @closing_remarks = 'あと少し！頑張ってね！'
  elsif judge_hour == 17
    @greeting = "お疲れ様！\n今日は " + now
    @closing_remarks = '今日も一日お疲れ様！'
  end

  puts @greeting

end

def send_mail(excel, file)

  Mail.defaults do
    delivery_method :smtp, {
      :address => 'sample',
      :port => 25,
      :domain => 'sample',
      :user_name => "#{mail_from}",
      :password => "#{mail_passwd}",
      :authentication => :login,
      :enable_starttls_auto => true
    }
  end

  mail = Mail.new do
    from    'sample'
    to      'sample'
    subject 'subject text'
    body     File.read("./body.txt")
    add_file './sample/xlxs'
  end
  
end

def update_clock_out(clock_out)

    tmp_clock_out_cell = cell.Address.to_s
    clock_out_cell = sheet.range(tmp_clock_out_cell.gsub(/A/, 'D'))
    target_cell_No = tmp_clock_out_cell.delete("^0-9").to_i

    while clock_out_cell.value == nil do      
      target_cell_No = target_cell_No - 1
      target_cell_address = '$D$' + target_cell_No.to_s
      clock_out_cell = sheet.range(target_cell_address)
    end

    clock_out_cell.value = clock_out

end

def update_clock_in(clock_in)

  tmp_clock_in_cell = cell.Address.to_s
  clock_in_cell = sheet.range(tmp_clock_in_cell.gsub(/A/, 'C'))
  clock_in_cell.value = clock_in

end

def update_excel(excel, file, sheet_num = 1, today)

  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)

  loop do
    puts '出退勤時刻を更新する？(y/n)'
    @answer_of_update = gets.chomp

    if @answer_of_update != "y" and @answer_of_update != "n"
      puts '「y」か「n」で入力してね'
    elsif  @answer_of_update == "y" or @answer_of_update == "n"
      break
    end
    
  end
  
  if @answer_of_update == 'y'
    puts '直近の退勤時間を入力してね！'
    tmp_clock_out = gets
    
    puts '今日の出勤時間を入力してね！'
    tmp_clock_in = gets
    
    clock_out = Time.parse(tmp_clock_out)
    clock_in = Time.parse(tmp_clock_in)
  end
  
  last_day = today.day - 1

  sheet.range('A10:A40').each do |cell|

    tmp_day = cell.value

    if @answer_of_update == 'y'
      if last_day == tmp_day.day
        update_clock_out(clock_out)
      end

      if today.day == tmp_day.day
        update_clock_in(clock_in)
      end
      puts '更新完了！'
    end 
  end
  # 保存
  book.saveAs(file)
  # ファイルを閉じる
  book.close
end

def main()

  today = Time.now()

  greeting(today)


  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  #file = fso.GetAbsolutePathName('./sample.xlsx')
  file = fso.GetAbsolutePathName('sample.xlsm')

  excel = init_excel()

  #send_mail(excel, file)
  update_excel(excel, file, today)

  excel.quit()

  puts @closing_remarks
  sleep(3)

end

main()

