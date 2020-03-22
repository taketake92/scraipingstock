require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

session = GoogleDrive::Session.from_config("config.json")

#まとめの結果シート
rid = "1shbIpdUhBLanoG6OVxX65flzBUqrYNPOilIm57Owzcc"
rws = session.spreadsheet_by_key(rid).worksheet_by_title("シート1")

p rws

count = 0
i = 1

#対象のシート
CSV.foreach("./SpreetheetsIdInfo.csv") do |line|

  p line[0]

  ws = session.spreadsheet_by_key(line[1]).worksheet_by_title("シート1")

  #今現在の最終行を取得
  for n in 3..200
    b = ws["A#{n}"]

    p b

    if b != "" then
      rws[n + count,0+i] = n + count - 2
      rws[n + count,1+i] = ws["A#{n}"]
    	rws[n + count,2+i] = ws["B#{n}"]
    	rws[n + count,3+i] = ws["C#{n}"]
    	rws[n + count,4+i] = ws["D#{n}"]
    	rws[n + count,5+i] = ws["E#{n}"]
    	rws[n + count,6+i] = ws["F#{n}"]
    	rws[n + count,7+i] = ws["G#{n}"]
    	rws[n + count,8+i] = ws["H#{n}"]
    	rws[n + count,9+i] = ws["I#{n}"]
    	rws[n + count,10+i] = ws["J#{n}"]
    	rws[n + count,11+i] = ws["K#{n}"]
    	rws[n + count,12+i] = ws["L#{n}"]
    	rws[n + count,13+i] = ws["M#{n}"]
    	rws[n + count,14+i] = ws["N#{n}"]
    	rws[n + count,15+i] = ws["O#{n}"]
    	rws[n + count,16+i] = ws["P#{n}"]
    	rws[n + count,17+i] = ws["Q#{n}"]
    	rws[n + count,18+i] = ws["R#{n}"]
    	rws[n + count,19+i] = ws["S#{n}"]
    	rws[n + count,20+i] = ws["T#{n}"]
    	rws[n + count,21+i] = ws["U#{n}"]
    	rws[n + count,22+i] = ws["V#{n}"]
    	rws[n + count,23+i] = ws["W#{n}"]
    	rws[n + count,24+i] = ws["X#{n}"]
    	rws[n + count,25+i] = ws["Y#{n}"]
    	rws[n + count,26+i] = ws["Z#{n}"]
    	rws[n + count,27+i] = ws["AA#{n}"]
    	rws[n + count,28+i] = ws["AB#{n}"]
    	rws[n + count,29+i] = ws["AC#{n}"]
    	rws[n + count,30+i] = ws["AD#{n}"]
    	rws[n + count,31+i] = ws["AE#{n}"]
    	rws[n + count,32+i] = ws["AF#{n}"]
    	rws[n + count,33+i] = ws["AG#{n}"]
    	rws[n + count,34+i] = ws["AH#{n}"]
    	rws[n + count,35+i] = ws["AI#{n}"]
    	rws[n + count,36+i] = ws["AJ#{n}"]
    	rws[n + count,37+i] = ws["AK#{n}"]
    	rws[n + count,38+i] = ws["AL#{n}"]
    	rws[n + count,39+i] = ws["AM#{n}"]
    	rws[n + count,40+i] = ws["AN#{n}"]
    	rws[n + count,41+i] = ws["AO#{n}"]
    	rws[n + count,42+i] = ws["AP#{n}"]
    	rws[n + count,43+i] = ws["AQ#{n}"]
    	rws[n + count,44+i] = ws["AR#{n}"]
    	rws[n + count,45+i] = ws["AS#{n}"]
    	rws[n + count,46+i] = ws["AT#{n}"]
    	rws[n + count,47+i] = ws["AU#{n}"]
    	rws[n + count,48+i] = ws["AV#{n}"]
    	rws[n + count,49+i] = ws["AW#{n}"]
      rws[n + count,50+i] = ws["AX#{n}"]
      rws[n + count,51+i] = ws["AY#{n}"]
      rws[n + count,52+i] = ws["AZ#{n}"]
      rws[n + count,53+i] = ws["BA#{n}"]
      rws[n + count,54+i] = ws["BB#{n}"]
      rws[n + count,55+i] = ws["BC#{n}"]

    else
      count = n + count - 3
      p count
      break
    end

  end

  sleep(rand(3..8.7))
  rws.save
end
