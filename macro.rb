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

def find_target_cell(cell, purpose)
  case purpose
  when 'Update' then
    purpose = '$D$'
  when 'SendMail' then
    purpose = '$A$'
  end

  tmp_target_cell_address = cell.Address.to_s
  target_cell = sheet.range(tmp_target_cell_address.gsub(/A/, 'D'))
  target_cell_address_No = tmp_target_cell.delete("^0-9").to_i

  while target_cell.value == nil do      
    target_cell_address_No = target_cell_address_No - 1
    target_cell_address = purpose + target_cell_address_No.to_s
    target_cell = sheet.range(target_cell_address)
  end
  return target_cell 
end

def update_clock_in(clock_in)
  tmp_clock_in_cell = cell.Address.to_s
  clock_in_cell = sheet.range(tmp_clock_in_cell.gsub(/A/, 'C'))
  clock_in_cell.value = clock_in
end

def validate_input(question)
  loop do
    puts question
    answer = gets.chomp

    if answer != "y" and answer != "n"
      puts '「y」か「n」で入力してね'
    elsif answer == "y" or answer == "n"
      return　answer
      break
    end
  end
end

def update_excel(excel, file, sheet_num = 1, today)
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(sheet_num)

  work_time_question = '出退勤時刻を更新する？(y/n)'
  answer_of_update = validate_input(work_time_question)
  
  if answer_of_update == 'y'
    puts '直近の退勤時間を入力してね！'
    tmp_clock_out = gets
    puts '今日の出勤時間を入力してね！'
    tmp_clock_in = gets
    
    clock_out = Time.parse(tmp_clock_out)
    clock_in = Time.parse(tmp_clock_in)
  end
  
  last_day = today.day - 1
  end_of_month = Date.new(today.year, today.month, -1)

  sheet.range('A10:A40').each do |cell|
    tmp_day = cell.value

    if answer_of_update == 'y'
      if last_day == tmp_day.day
        purpose = 'Update'
        clock_out_cell = find_target_cell(cell, purpose)
        clock_out_cell.value = clock_out
      end

      if today.day == tmp_day.day
        update_clock_in(clock_in)
      end
      puts '更新完了！'
    end

    case tmp_day.day
    when 18 then
      purpose = 'SendMail'
      @to_bright_day = find_target_cell(cell, purpose)
    when end_of_month.day then
      purpose = 'SendMail'
      @to_me_day = find_target_cell(cell, purpose)
    end
  end
  # 保存
  book.saveAs(file)
  # ファイルを閉じる
  book.close
end

def sendmail_decision(today)
  case today.day
  when @to_bright_day.day then
    puts '今日は現場勤務表提出日だよ！勤務表の中身を確認してね！'
    question = '今すぐ自社にメールを送る？(y/n)'
    answer_of_sendmail = validate_input(question)
    if answer_of_sendmail == 'y'
      destination = 'bright'
      send_mail(destination, today)
    elsif answer_of_sendmail == 'n'
      puts '今日中に送ってね！'
    end
  when @to_me_day.day then
    puts '今日は月末だよ！勤務表の中身を確認してね！'
    question = '今すぐ自分宛てにメールを送る？(y/n)'
    answer_of_sendmail = validate_input(question)
    if answer_of_sendmail == 'y'
      destination = 'me'
      send_mail(destination, today)
    elsif answer_of_sendmail == 'n'
      puts '今日中に送ってね！'
    end
  end
end

def send_mail(destination, today)

  mail_info = mail_creation(destination, today)

  puts "メールの内容を確認してね！
        差出人：#{mail_info[from]}
        宛先：#{mail_info[to]}
        cc：#{mail_info[cc]}
        件名：#{mail_info[subject]}
        本文：#{mail_info[body]}"
  question = 'この内容で送ってもいい？(y/n)'
  sendmail_decision = validate_input(question)
  
  if sendmail_decision == 'y'
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
      from     "#{mail_info[from]}"
      to       "#{mail_info[to]}"
      cc       "#{mail_info[cc]}"
      subject  "#{mail_info[subject]}"
      body     "#{mail_info[body]}"
      add_file ""
    end
    mail.deliver
    puts '送信完了！'
  elsif sendmail_decision == 'n'
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
                  body:     '各位

                  お疲れ様です。です。
                  今月分の現場勤務表を送付致します。
                  ご確認よろしくお願いいたします。
                  '
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

  greeting(today)

  # OLE32用FileSystemObject生成
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  #file = fso.GetAbsolutePathName('./sample.xlsx')
  file = fso.GetAbsolutePathName('sample.xlsm')

  excel = init_excel()
  update_excel(excel, file, today)
  excel.quit()

  sendmail_decision(today) 

  puts @closing_remarks
  sleep(3)
end

main()
