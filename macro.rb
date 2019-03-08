#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
STDOUT.sync = true

require 'win32ole'
require 'time'
require 'date'
#require 'mail'

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
    opening_remarks = "おはよう！\n今日は " + now
    @closing_remarks = '今日も一日頑張ろう！'
  elsif judge_hour >=13 and judge_hour <= 16
    opening_remarks = "こんにちは！\n今日は " + now
    @closing_remarks = 'あと少し！頑張ってね！'
  elsif judge_hour >= 17
    opening_remarks = "お疲れ様！\n今日は " + now
    @closing_remarks = '今日も一日お疲れ様！'
  end
  puts opening_remarks
end

def find_target_cell(sheet, cell, purpose)
  tmp_target_cell_address = cell.Address.to_s
  target_cell = sheet.range(tmp_target_cell_address.gsub(/A/, 'D'))
  target_cell_address_No = tmp_target_cell_address.delete("^0-9").to_i

  case purpose
  when 'Update'
    purpose = '$D$'
  when 'SendMail'
    purpose = '$A$'
    if target_cell.value != nil
      target_cell = sheet.range(tmp_target_cell_address)
    end
  end

  while target_cell.value == nil do      
    target_cell_address_No = target_cell_address_No - 1
    target_cell_address = purpose + target_cell_address_No.to_s
    target_cell = sheet.range(target_cell_address)
  end
  return target_cell 
end

def update_clock_in(sheet, cell, clock_in)
  tmp_clock_in_cell = cell.Address.to_s
  clock_in_cell = sheet.range(tmp_clock_in_cell.gsub(/A/, 'C'))
  clock_in_cell.value = clock_in
end

def validate_input(question)
  loop do
    puts question
    @answer = gets.chomp

    if @answer != "y" and @answer != "n"
      puts '「y」か「n」で入力してね'
    elsif @answer == "y" or @answer == "n"
      break
    end
  end
end

def update_excel(excel, file, sheet_num = 1, today)
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)

  work_time_question = '出退勤時刻を更新する？(y/n)'
  validate_input(work_time_question)
  
  if @answer == 'y'
    puts '直近の退勤時間を入力してね！'
    clock_out = gets.chomp
    puts '今日の出勤時間を入力してね！'
    clock_in = gets.chomp

  end
  
  last_day = today.day - 1
  end_of_month = Date.new(today.year, today.month, -1)

  sheet.range('A10:A40').each do |cell|
    tmp_day = cell.value

    if @answer == 'y'
      if last_day == tmp_day.day
        purpose = 'Update'
        clock_out_cell = find_target_cell(sheet, cell, purpose)
        clock_out_cell.value = clock_out
      end

      if today.day == tmp_day.day
        update_clock_in(sheet, cell, clock_in)
      end
      @result_report = '更新完了！'
    end

    if today.day == 18 and tmp_day.day == 18
      purpose = 'SendMail'
      target_cell = find_target_cell(sheet, cell, purpose)
      @to_bright_day = target_cell.value
    elsif today.day == end_of_month.day and tmp_day.day == end_of_month.day
      purpose = 'SendMail'
      target_cell = find_target_cell(sheet, cell, purpose)
      @to_me_day = target_cell.value
    end
  end
  # 保存
  book.saveAs(file)
  # ファイルを閉じる
  book.close
  puts @result_report
end

def sendmail_decision(today)
  case today.day
  when @to_bright_day.day
    puts '今日は現場勤務表提出日だよ！勤務表の中身を確認してね！'
    question = '今すぐ自社にメールを送る？(y/n)'
    validate_input(question)
    if @answer == 'y'
      destination = 'bright'
      send_mail(destination, today)
    elsif @answer == 'n'
      puts '今日中に送ってね！'
    end
  when @to_me_day.day
    puts '今日は月末だよ！勤務表の中身を確認してね！'
    question = '今すぐ自分宛てにメールを送る？(y/n)'
    @answer = validate_input(question)
    if @answer == 'y'
      destination = 'me'
      send_mail(destination, today)
    elsif @answer == 'n'
      puts '今日中に送ってね！'
    end
  end
end

def send_mail(destination, today)
  mail_info = mail_creation(destination, today)

  puts "メールの内容を確認してね！\n差出人：#{mail_info[:from]}\n宛先：#{mail_info[:to]}\
  \ncc：#{mail_info[:cc]}\n件名：#{mail_info[:subject]}\n本文：\n\n#{mail_info[:body]}"
  question = 'この内容で送ってもいい？(y/n)'
  sendmail_decision = validate_input(question)
  
  if @answer == 'y'
    #Mail.defaults do
    #  delivery_method :smtp, {
    #    :address => 'sample',
    #    :port => 25,
    #    :domain => 'sample',
    #    :user_name => "#{mail_from}",
    #    :password => "#{mail_passwd}",
    #    :authentication => :login,
    #    :enable_starttls_auto => true
    #  }
    #end
#
    #mail = Mail.new do
    #  from     "#{mail_info[:from]}"
    #  to       "#{mail_info[:to]}"
    #  cc       "#{mail_info[:cc]}"
    #  subject  "#{mail_info[:subject]}"
    #  body     "#{mail_info[:body]}"
    #  add_file ""
    #end
    #mail.deliver
    puts '送信完了！'
  elsif @answer == 'n'
    puts '送信をキャンセルしたよ！'
  end
end

def mail_creation(destination, today)
  subject = today.strftime('_勤務表 %Y年%m月分')
  case destination
  when 'bright'
    mail_info = {
                  from:     '',
                  to:       '',
                  cc:       '',
                  subject:  '',
                  body:     "各位\n\nお疲れ様です。です。\n今月分の現場勤務表を送付致します。\nご確認よろしくお願いいたします。\n\n"
                }
  when 'me'
    mail_info = {
                  from:     '',
                  to:       '',
                  cc:       '',
                  subject:  '',
                  body:     '今日中に自社勤務表を宛にメールしてね！'
                }
  end
  return mail_info
end

def main()
  today = Time.now
  today = Time.new(today.year, today.month, 18, 9, 30)
  #end_of_month = Date.new(today.year, today.month, -1)
  #today = end_of_month

  greeting(today)

  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  #file = fso.GetAbsolutePathName('./sample.xlsx')
  file = fso.GetAbsolutePathName('C:\Users\HMP01156\OUT\tmp.xlsm')

  excel = init_excel()
  update_excel(excel, file, today)
  excel.quit()

  if @to_bright_day != nil or @to_me_day != nil
    sendmail_decision(today)
  end

  puts @closing_remarks
  sleep(3)
end

main()
