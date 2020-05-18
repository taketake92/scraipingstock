require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

mCodes = [
'0000', #日経平均
'0010', #TOPIX
'0012', #マザーズ指数
'0100', #日系JASDAQ
'0800', #NYダウ
'0802', #NASDAQ
1332,
1333,
1357,
1447,
1570,
1605,
1719,
1720,
1721,
1801,
1802,
1803,
1808,
1812,
1820,
1860,
1883,
1890,
1893,
1925,
1928,
1942,
1944,
1951,
1963,
2002,
2060,
2146,
2160,
2162,
2191,
2269,
2282,
2362,
2370,
2413,
2432,
2433,
2471,
2501,
2502,
2503,
2531,
2685,
2768,
2801,
2802,
2871,
2930,
3003,
3038,
3064,
3086,
3092,
3099,
3101,
3103,
3105,
3186,
3254,
3258,
3289,
3323,
3356,
3382,
3401,
3402,
3405,
3407,
3436,
3604,
3665,
3776,
3861,
3863,
3990,
4004,
4005,
4021,
4042,
4043,
4061,
4063,
4151,
4183,
4188,
4204,
4205,
4208,
4217,
4272,
4321,
4324,
4344,
4347,
4364,
4452,
4502,
4503,
4506,
4507,
4519,
4523,
4528,
4543,
4563,
4564,
4565,
4568,
4572,
4578,
4583,
4587,
4588,
4592,
4631,
4661,
4666,
4676,
4681,
4689,
4704,
4714,
4739,
4751,
4755,
4813,
4901,
4902,
4911,
4974,
4978,
5019,
5020,
5101,
5108,
5201,
5202,
5214,
5233,
5301,
5332,
5333,
5401,
5406,
5411,
5631,
5703,
5706,
5711,
5713,
5714,
5727,
5801,
5802,
5803,
5901,
5912,
6035,
6062,
6095,
6096,
6098,
6101,
6113,
6178,
6184,
6235,
6301,
6302,
6305,
6326,
6361,
6367,
6395,
6460,
6471,
6472,
6473,
6479,
6501,
6502,
6503,
6504,
6506,
6569,
6645,
6674,
6701,
6702,
6703,
6724,
6752,
6754,
6758,
6762,
6770,
6838,
6841,
6857,
6861,
6890,
6902,
6920,
6952,
6954,
6963,
6965,
6971,
6976,
6981,
6988,
7003,
7004,
7011,
7012,
7013,
7148,
7172,
7186,
7198,
7201,
7202,
7203,
7205,
7211,
7224,
7238,
7261,
7267,
7269,
7270,
7272,
7518,
7550,
7564,
7731,
7733,
7735,
7746,
7751,
7752,
7762,
7777,
7816,
7832,
7836,
7911,
7912,
7936,
7951,
7974,
7980,
8001,
8002,
8015,
8028,
8031,
8035,
8053,
8056,
8058,
8111,
8136,
8226,
8233,
8252,
8253,
8267,
8303,
8304,
8306,
8308,
8309,
8316,
8331,
8354,
8355,
8411,
8601,
8604,
8628,
8630,
8725,
8729,
8750,
8766,
8795,
8801,
8802,
8804,
8830,
8893,
8897,
8934,
9001,
9005,
9006,
9007,
9008,
9009,
9020,
9021,
9022,
9042,
9062,
9064,
9101,
9104,
9107,
9201,
9202,
9263,
9301,
9401,
9404,
9412,
9424,
9432,
9433,
9434,
9437,
9501,
9502,
9503,
9531,
9532,
9602,
9603,
9613,
9616,
9681,
9706,
9716,
9735,
9766,
9783,
9792,
9983,
9984,

]



session = GoogleDrive::Session.from_config("config.json")

rsRow = 3

id = "1dWGVF8kMglGzRpHoAfOM_3l5BAYJX0baJAb2cW_tgnk"
rs = session.spreadsheet_by_key(id).worksheet_by_title("シート1")

mCodes.each{|mCode|

  #テンプレートからコピーして新規にsheet作成
  #テンプレート読み込み
  templeate_sheets = ""
  copy_sheets = ""
  templeate_sheets = session.spreadsheet_by_key("1nEV6YwvaoR94XjWZADDKla-GZEifqtF1RewZ6ali1Io")
  #テンプレート作成
  sub_folder_id = '1FqYOQLzAfR59Bb6mZ4hgV6ZKDo8Bo547'



  sub_folder = session.collection_by_id(sub_folder_id)

  p mCode
  copy_sheets = templeate_sheets.copy("#{mCode}test")
  sub_folder.add(copy_sheets)

  #テンプレートのid取得
  sheet_id = copy_sheets.id
  #テンプレートでシートインスタンス
  sheets = session.spreadsheet_by_key(sheet_id).worksheets[1]

   p sheets;

  row = 3
  count = File.read("./#{mCode}.csv").count("\n")
  CSV.foreach("./#{mCode}.csv") do |line|

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
  	sheets[row,12] = line[8].delete!("株") #出来高
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



  test = CSV.open("./SpreetheetsIdInfo.csv","a") do |csv|
  csv << [
        mCode,
        sheet_id,
        ]
  end

  ws = session.spreadsheet_by_key(sheet_id).worksheet_by_title("シート0")

  rs[rsRow,1] = ws['A1']
  rs[rsRow,2] = ws['B1']
  # rs[rsRow,7] = ws['C1']
  rs[rsRow,8] = ws['D1']
  rs[rsRow,9] = ws['E1']
  rs[rsRow,10] = ws['F1']
  rs[rsRow,11] = ws['G1']
  rs[rsRow,12] = ws['H1']
  rs[rsRow,13] = ws['I1']
  rs[rsRow,14] = ws['J1']
  rs[rsRow,15] = ws['K1']
  rs[rsRow,16] = ws['L1']
  rs[rsRow,17] = ws['M1']
  rs[rsRow,18] = ws['N1']
  rs[rsRow,19] = ws['O1']
  rs[rsRow,20] = ws['P1']
  rs[rsRow,21] = ws['Q1']
  # rs[rsRow,24] = ws['R1']
  # rs[rsRow,25] = ws['S1']
  # rs[rsRow,26] = ws['T1']
  rs[rsRow,24] = ws['U1']
  rs[rsRow,25] = ws['V1']
  rs[rsRow,26] = ws['W1']
  rs[rsRow,27] = ws['X1']
  rs[rsRow,28] = ws['Y1']
  rs[rsRow,29] = ws['Z1']
  rs[rsRow,30] = ws['AA1']
  rs[rsRow,31] = ws['AB1']
  rs[rsRow,32] = ws['AC1']
  rs[rsRow,33] = ws['AD1']
  rs[rsRow,34] = ws['AE1']

  rs[rsRow,35] = ws['B8']

  rs[rsRow,36] = ws['AF1']
  rs[rsRow,37] = ws['AG1']
  rs[rsRow,38] = ws['AH1']
  rs[rsRow,39] = ws['AI1']
  rs[rsRow,40] = ws['AJ1']
  rs[rsRow,41] = ws['AK1']
  rs[rsRow,42] = ws['AL1']
  rs[rsRow,43] = ws['AM1']
  rs[rsRow,44] = ws['AN1']
  rs[rsRow,45] = ws['AO1']
  rs[rsRow,46] = ws['AP1']
  rs[rsRow,47] = ws['AQ1']
  rs[rsRow,48] = ws['AR1']
  rs[rsRow,49] = ws['AS1']
  rs[rsRow,50] = ws['AT1']
  rs[rsRow,51] = ws['AU1']
  rs[rsRow,52] = ws['AV1']
  rs[rsRow,53] = ws['AW1']
  rs[rsRow,54] = ws['AX1']
  rs[rsRow,55] = ws['AY1']
  rs[rsRow,56] = ws['AZ1']
  rs[rsRow,57] = ws['BA1']
  rs[rsRow,58] = ws['BB1']
  rs[rsRow,59] = ws['BC1']
  rs[rsRow,60] = ws['BD1']
  rs[rsRow,61] = ws['BE1']
  rs[rsRow,62] = ws['BF1']
  rs[rsRow,63] = ws['BG1']
  rs[rsRow,64] = ws['BH1']
  rs[rsRow,65] = ws['BI1']
  rs[rsRow,66] = ws['BJ1']
  rs[rsRow,67] = ws['BK1']
  rs[rsRow,68] = ws['BL1']
  rs[rsRow,69] = ws['BM1']
  rs[rsRow,70] = ws['BN1']

  # 乖離線
  rs[rsRow,71] = ws['A5']
  rs[rsRow,72] = ws['B5']
  rs[rsRow,73] = ws['C5']
  rs[rsRow,74] = ws['D5']
  rs[rsRow,75] = ws['E5']
  rs[rsRow,76] = ws['F5']
  rs[rsRow,77] = ws['G5']
  rs[rsRow,78] = ws['H5']
  rs[rsRow,79] = ws['I5']

  rs[rsRow,80] = ws['B7']
  rs[rsRow,81] = ws['C7']
  rs[rsRow,82] = ws['D7']
  rs[rsRow,83] = ws['E7']
  rs[rsRow,84] = ws['F7']

  rs[rsRow,85] = sheet_id


  rsRow+= 1

  rs.save

  # sleep(rand(2..3))

}


#file_info.num_rows
#ファイルの最終行の行番号取得
