require 'rubyXL/convenience_methods/cell'
require 'rubyXL/convenience_methods/workbook'
require 'rubyXL'

mCodes = [



'0000',
2002,
2269,
2282,
2501,
2502,
2503,
2531,
2801,
2802,
2871,
2914,
3101,
3103,
3401,
3402,
3861,
3863,
3405,
3407,
4004,
4005,
4021,
4042,
4043,
4061,
4063,
4183,
4188,
4208,
4272,
4452,
4631,
4901,
4911,
4151,
4502,
4503,
4506,
4507,
4519,
4523,
4568,
4578,
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
3436,
5703,
5706,
5711,
5713,
5714,
5801,
5802,
5803,
5901,
5631,
6113,
6301,
6302,
6305,
6326,
6361,
6367,
6471,
6472,
6473,
7004,
7011,
7013,
3105,
6479,
6501,
6503,
6504,
6506,
6645,
6674,
6701,
6702,
6703,
6724,
6752,
6758,
6762,
6770,
6841,
6857,
6902,
6952,
6954,
6971,
6976,
6988,
7735,
7751,
7752,
8035,
7003,
7012,
7201,
7202,
7203,
7205,
7211,
7261,
7267,
7269,
7270,
7272,
4543,
4902,
7731,
7733,
7762,
7832,
7911,
7912,
7951,
1332,
1333,
1605,
1721,
1801,
1802,
1803,
1808,
1812,
1925,
1928,
1963,
2768,
8001,
8002,
8015,
8031,
8053,
8058,
3086,
3099,
3382,
8028,
8233,
8252,
8267,
9983,
7186,
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
8253,
3289,
8801,
8802,
8804,
8830,
9001,
9005,
9007,
9008,
9009,
9020,
9021,
9022,
9062,
9064,
9101,
9104,
9107,
9202,
9301,
9412,
9432,
9433,
9437,
9613,
9984,
9501,
9502,
9503,
9531,
9532,
2432,
4324,
4689,
4704,
4751,
4755,
6098,
6178,
9602,
9681,
9735,
9766,
3990,
3665,
6184,
6754,
6963,
7974,
3038,
3064,
6920,
7550,
7564,
7816,
1357,
4588,
4344,
3356,
9434,
3092,
6861,
1570,
3604,
2162,
8893,
3323,
2060,
6569,
1890,
9277,
8897,
7172,
2362,
7198,
2146,
7148,
3254,
4528,
2370,
4564,
4587,
4974,
4592,
4565,
4583,
4572,
2191,
4978,
6981,
4217,
4563,
5727,
4205,
6965,
7238,
6062,
9263,
7980,
7777,
6502,
7936,
8111,
1719,
1820,
1860,
1883,
1893,
5912,
6395,
7224,
1951,
2413,
3776,
4739,
4813,
6101,
6235,
6838,
6890,
7518,
7746,
8056,
8226,
9424,
1447,
1719,
1720,
1820,
1860,
1883,
1893,
1942,
1944,
2433,
3003,
3258,
4204,
4321,
4661,
4666,
4676,
4681,
4714,
5912,
6395,
6460,
7836,
7936,
8111,
8136,
8934,
9006,
9042,
9201,
9401,
9404,
9603,
9616,
9706,
9716,
9783,
9792,
4347,
2160,
6096,
6035,
4364,
2471,
2685,
2930,
6095,
3186,

]

rbook = RubyXL::Parser.parse("./allStockList.xlsx")
rbook.calc_pr.full_calc_on_load = true
rbook.calc_pr.calc_completed = true
rbook.calc_pr.calc_on_save = true
rbook.calc_pr.force_full_calc = true
rsheet = rbook['シート1']

mCodes.each{|mCode|

book = RubyXL::Parser.parse("./sample.xlsx")
book.calc_pr.full_calc_on_load = true
book.calc_pr.calc_completed = true
book.calc_pr.calc_on_save = true
book.calc_pr.force_full_calc = true
sheet = book['シート1']

row = 2
i = 1
  CSV.foreach("./../#{mCode}.csv") do |line|

	sheet.add_cell(row,1-i,line[0]) #銘柄コード
  	sheet.add_cell(row,2-i,line[1]) #日付
  	sheet.add_cell(row,3-i,line[10]) #前日比
  	sheet.add_cell(row,4-i,line[11]) #前日比(㌫)
  	sheet.add_cell(row,5-i,line[2].to_i) #始値
  	sheet.add_cell(row,7-i,line[3].to_i) #高値
  	sheet.add_cell(row,8-i,line[4]) #高値(時刻)
  	sheet.add_cell(row,9-i,line[5].to_i) #安値
  	sheet.add_cell(row,10-i,line[6]) #安値(時刻)
  	sheet.add_cell(row,11-i,line[7].to_i) #終値
  	sheet.add_cell(row,12-i,line[8]) #出来高
  	sheet.add_cell(row,13-i,line[9]) #売買代金
  	sheet.add_cell(row,14-i,line[13]) #約定回数
  	sheet.add_cell(row,15-i,line[14]) #時価総額
  	row += 1

end

book.write "./stockList/#{mCode}.xlsx"

sleep(5)

#再度開いてシートの結果を取得
rbook = RubyXL::Parser.parse("./stockList/#{mCode}.xlsx")
rbook.calc_pr.full_calc_on_load = true
rbook.calc_pr.calc_completed = true
rbook.calc_pr.calc_on_save = true
rbook.calc_pr.force_full_calc = true
rsheet = book['シート1']

rsheet.add_cell(rrow,sheet)

}
