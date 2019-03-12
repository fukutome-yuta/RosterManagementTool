#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
STDOUT.sync = true

require 'win32ole'
require 'time'
require 'date'
require 'mail'

TODAY = Time.now

# Excel VBA定数のロード
module Excel; end

def main()
  greeting()

  fso = WIN32OLE.new('Scripting.FileSystemObject')
  file = fso.GetAbsolutePathName('sample.xlsm')

  excel = init_excel()
  update_excel(excel, file)
  excel.quit()

  sendmail_decision()
  
  puts @closing_remarks
  sleep(3)
end

def greeting()
  now = TODAY.strftime('%Y年 %m月 %d日 (%a)')
  judge_hour = TODAY.hour

  #現在時刻で挨拶の内容を変更する
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

def init_excel()
  # Excelオブジェクト生成
  excel = WIN32OLE.new('Excel.Application')
  excel.visible = false
  # 上書きメッセージを抑制
  excel.displayAlerts = false

  WIN32OLE.const_load(excel, Excel)

  return excel
end

#エクセルの更新とメール送信判断に使う日付を取得する
def update_excel(excel, file)
  worktime_question = '出退勤時刻を更新する？(y/n)'
  update_worktime = validate_input(worktime_question)
  if update_worktime
    puts '直近の退勤時間を入力してね！'
    clock_out = gets.chomp
    puts '今日の出勤時間を入力してね！'
    clock_in = gets.chomp
  end

  holidays_question = '休みの予定を更新する？(y/n)'
  update_holidays = validate_input(holidays_question)
  if update_holidays
    loop do
      puts '休む予定の日付を[1～31]の数字で入力してね！'
      get_day = gets.chomp
      if get_day !~ /[1-31]/
        puts '[1～31]の整数で入力してね！'
      else
        @holidays = get_day.to_i
        break
      end
    end
    loop do
      puts "休暇事由を[1～6]の数字で入力してね！\n[年休:1, 欠勤:2, 公休:3, 振替休:4, 祝日:5, 休業日:6]"
      get_reason = gets.chomp
      if get_reason !~ /[1-6]/
        puts '[1～6]の整数で入力してね！'
      else
        case get_reason
        when "1"
          @holiday_reason = "年休"
        when "2"
          @holiday_reason = "欠勤"
        when "3"
          @holiday_reason = "公休"
        when "4"
          @holiday_reason = "振替休"
        when "5"
          @holiday_reason = "祝日"
        when "6"
          @holiday_reason = "休業日"
        end
        break
      end
    end
  end
  
  book = excel.Workbooks.Open(file)
  sheet = book.Worksheets(1)
  last_day = TODAY.day - 1

  sheet.range('A10:A40').each do |cell|
    index_day = cell.value

    if update_worktime
      if last_day == index_day.day
        purpose = 'Update'
        clock_out_cell = find_target_cell(sheet, cell, purpose)
        clock_out_cell.value = clock_out
      end

      if TODAY.day == index_day.day
        update_clock_in(sheet, cell, clock_in)
      end
      @result_report = '更新完了！'
    end

    if update_holidays
      if @holidays == index_day.day
        holidays_cell_address = cell.Address.to_s
        holiday_of_clock_in_cell = sheet.range(holidays_cell_address.gsub(/A/, 'C'))
        holiday_of_clock_out_cell = sheet.range(holidays_cell_address.gsub(/A/, 'D'))
        holiday_reason_cell = sheet.range(holidays_cell_address.gsub(/A/, 'I'))
        holiday_of_clock_in_cell.value = ""
        holiday_of_clock_out_cell.value = ""
        holiday_reason_cell.value = @holiday_reason
      end
      @result_report = '更新完了！'
    end

    if TODAY.day == 18 and index_day.day == 18
      purpose = 'SendMail'
      target_cell = find_target_cell(sheet, cell, purpose)
      @to_mycompany = target_cell.value
    end
  end

  purpose = 'SendMail'
  end_of_month_cell = sheet.range('A40')
  last_working_day = find_target_cell(sheet, end_of_month_cell, purpose)
  @to_me = last_working_day.value
  # 保存
  book.saveAs(file)
  # ファイルを閉じる
  book.close
  puts @result_report
end

#入力された文字を検証
def validate_input(question)
  loop do
    puts question
    answer = gets.chomp

    if answer == "y"
      break true
    elsif answer == 'n'
      break false
    else
      puts '[y]か[n]で入力してね'
    end
  end
end

#update, sendmail 判断対象のセルを特定する
def find_target_cell(sheet, cell, purpose)
  indicator_cell_address = cell.Address.to_s  
  indicator_cell = sheet.range(indicator_cell_address.gsub(/A/, 'D'))
  #セルの数字のみを切り出し while のイテレータとして利用する
  target_cell_address_No = indicator_cell_address.delete("^0-9").to_i
  target_cell = indicator_cell
  
  #D列（退勤時刻）をさかのぼり、直近で入力のあるセルを特定する
  while target_cell.value == nil do      
    target_cell_address_No = target_cell_address_No - 1
    target_cell_address = '$D$' + target_cell_address_No.to_s
    target_cell = sheet.range(target_cell_address)
  end
  
  #送信日の判断のためA列（日付）のセルを返す
  if purpose == 'SendMail'
    if indicator_cell.value != nil
      target_cell = sheet.range(indicator_cell_address)
    else
      target_cell = sheet.range(target_cell_address.gsub(/D/, 'A'))
    end
  end
  return target_cell 
end

#出勤時刻の更新
def update_clock_in(sheet, cell, clock_in)
  tmp_clock_in_cell = cell.Address.to_s
  clock_in_cell = sheet.range(tmp_clock_in_cell.gsub(/A/, 'C'))
  clock_in_cell.value = clock_in
end

#現場勤務表送付日、自社勤務表送付日判断
def sendmail_decision()
  if @to_mycompany != nil
    if TODAY.day == @to_mycompany.day
      puts '今日は現場勤務表提出日だよ！勤務表の中身を確認してね！'
      question = '今すぐ自社宛てにメールを送る？(y/n)'
      send_to_mycompany = validate_input(question)
      if send_to_mycompany
        destination = 'to_mycompany'
        send_mail(destination)
      else
        puts "送信を見送るよ！\nあとで確認してから必ず今日中に送ってね！"
      end
    end
  else
    if TODAY.day == @to_me.day
      puts '今日は月末だよ！勤務表の中身を確認してね！'
      question = '今すぐ自分宛てにメールを送る？(y/n)'
      send_to_me = validate_input(question)
      if send_to_me
        destination = 'me'
        send_mail(destination)
      else
        puts "送信を見送るよ！\nあとで確認してから必ず今日中に送ってね！"
      end
    end
  end
end

#メールの作成 → 送信
def send_mail(destination)
  puts 'メールを作成するよ！(to_' + destination + ')'
  sleep(2)
  mail_info = mail_creation(destination)

  puts "メールの内容を確認してね！\n差出人：#{mail_info[:from]}\n宛先：#{mail_info[:to]}\
  \ncc：#{mail_info[:cc]}\n件名：#{mail_info[:subject]}\n本文：\n\n#{mail_info[:body]}"
  question = 'この内容で送ってもいい？(y/n)'
  sendmail = validate_input(question)
  
  if sendmail
    mail = Mail.new do
      from     "#{mail_info[:from]}"
      to       "#{mail_info[:to]}"
      cc       "#{mail_info[:cc]}"
      subject  "#{mail_info[:subject]}"
      body     "#{mail_info[:body]}"
      add_file 'sample.xlsm'
    end
    mail.deliver
    puts '送信完了！'
  else
    puts "送信をキャンセルしたよ！\nあとで確認してから必ず今日中に送ってね！"
  end
end

#現場、自社宛でメールの内容を変える
def mail_creation(destination)
    subject = TODAY.strftime('_現場勤務表 %Y年%m月分')
    case destination
    when 'to_mycompany'
      mail_info = {
        from:     '',
        to:       '',
        cc:       '',
        subject:  subject,
        body:     "各位\n\nお疲れ様です。\n今月分の現場勤務表を送付致します。\nご確認よろしくお願いいたします。"
      }
    when 'me'
      mail_info = {
        from:     '',
        to:       '',
        cc:       '',
        subject:  subject,
        body:     "今日中に自社勤務表を送ってね！"
      }
    end
    return mail_info
  end

main()
