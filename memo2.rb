require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

session = GoogleDrive::Session.from_config("config.json")

rsRow = 3

id = "1dWGVF8kMglGzRpHoAfOM_3l5BAYJX0baJAb2cW_tgnk"
rs = session.spreadsheet_by_key(id).worksheet_by_title("シート1")

CSV.foreach("./SpreetheetsIdInfo.csv") do |data|

  p data[0]

  #テンプレートでシートインスタンス
  sheets = session.spreadsheet_by_key(data[1]).worksheets[1]

  p sheets

  row = 3
  count = File.read("./#{data[0]}.csv").count("\n")

  CSV.foreach("./#{data[0]}.csv") do |line|

  	sheets[row,1] = line[0] #銘柄コード
  	sheets[row,2] = line[1] #日付
  	sheets[row,3] = line[10] #前日比
  	sheets[row,4] = line[11] #前日比(㌫)
  	sheets[row,5] = line[2] #始値
  	sheets[row,7] = line[3] #高値
  	sheets[row,8] = line[4] #高値(時刻)
  	sheets[row,9] = line[5] #安値
  	sheets[row,10] = line[6] #安値(時刻)
  	sheets[row,11] = line[7] #終値
  	sheets[row,12] = line[8] #出来高
  	sheets[row,13] = line[9] #売買代金
  	sheets[row,14] = line[13] #約定回数
  	sheets[row,15] = line[14] #時価総額
  	row += 1

  end

  # # Dumps all cells.
  # (1..sheets.num_rows).each do |row|
  #   (1..sheets.num_cols).each do |col|
  #     p sheets[row, col]
  #   end
  # end
# sheets.insertRowAfter(5);


# p sheets[count,16]
# p sheets[count,17]
# p sheets[count,18]
# p sheets[count,19]

  sheets.save

  ws = session.spreadsheet_by_key(data[1]).worksheet_by_title("シート0")

    rs[rsRow,1] = ws['A1']
    rs[rsRow,2] = ws['B1']
    rs[rsRow,3] = ws['C1']
    rs[rsRow,4] = ws['D1']
    rs[rsRow,5] = ws['E1']
    rs[rsRow,6] = ws['F1']
    rs[rsRow,7] = ws['G1']
    rs[rsRow,8] = ws['H1']
    rs[rsRow,9] = ws['I1']
    rs[rsRow,10] = ws['J1']
    rs[rsRow,11] = ws['K1']
    rs[rsRow,12] = ws['L1']
    rs[rsRow,13] = ws['M1']
    rs[rsRow,14] = ws['N1']
    rs[rsRow,15] = ws['O1']
    rs[rsRow,16] = ws['P1']
    rs[rsRow,17] = ws['Q1']
    rs[rsRow,18] = ws['R1']
    rs[rsRow,19] = ws['S1']
    rs[rsRow,20] = ws['T1']
    rs[rsRow,21] = ws['U1']
    rs[rsRow,22] = ws['V1']
    rs[rsRow,23] = ws['W1']
    rs[rsRow,24] = ws['X1']
    rs[rsRow,25] = ws['Y1']
    rs[rsRow,26] = ws['Z1']
    rs[rsRow,27] = ws['AA1']
    rs[rsRow,28] = ws['AB1']
    rs[rsRow,29] = ws['AC1']
    rs[rsRow,30] = ws['AD1']
    rs[rsRow,31] = ws['AE1']
    rs[rsRow,32] = ws['AF1']
    rs[rsRow,33] = ws['AG1']
    rs[rsRow,34] = ws['AH1']
    rs[rsRow,35] = ws['AI1']
    rs[rsRow,36] = ws['AJ1']
    rs[rsRow,37] = ws['AK1']
    rs[rsRow,38] = ws['AL1']
    rs[rsRow,39] = ws['AM1']
    rs[rsRow,40] = ws['AN1']
    rs[rsRow,41] = ws['AO1']
    rs[rsRow,42] = ws['AP1']
    rs[rsRow,43] = ws['AQ1']
    rs[rsRow,44] = ws['AR1']
    rs[rsRow,45] = ws['AS1']
    rs[rsRow,46] = ws['AT1']
    rs[rsRow,47] = ws['AU1']
    rs[rsRow,48] = ws['AV1']
    rs[rsRow,49] = ws['AW1']
    rs[rsRow,50] = ws['AX1']
    rs[rsRow,51] = ws['AY1']
    rs[rsRow,52] = ws['AZ1']
    rs[rsRow,53] = ws['BA1']
    rs[rsRow,54] = ws['BB1']
    rs[rsRow,55] = ws['BC1']
    rs[rsRow,56] = ws['BD1']
    rs[rsRow,57] = ws['BE1']
    rs[rsRow,58] = ws['BF1']
    rs[rsRow,59] = ws['BG1']
    rs[rsRow,60] = ws['BH1']
    # 乖離線
    rs[rsRow,61] = ws['A5']
    rs[rsRow,62] = ws['B5']
    rs[rsRow,63] = ws['C5']
    rs[rsRow,64] = ws['D5']
    rs[rsRow,65] = ws['E5']
    rs[rsRow,66] = ws['F5']
    rs[rsRow,67] = ws['G5']
    rs[rsRow,68] = ws['H5']
    rs[rsRow,69] = ws['I5']
    

  rsRow+= 1

  rs.save
  
  c = rand(3..8.7)

  sleep(c)

end

#file_info.num_rows
#ファイルの最終行の行番号取得
