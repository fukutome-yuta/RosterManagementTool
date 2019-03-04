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