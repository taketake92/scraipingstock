#まとめの結果シート
rid = "11OV4wCVw_gaSGItgnraaRziuXwLTaOJADU1d3RQ7_bY"
rws = session.spreadsheet_by_key(rid).worksheet_by_title("シート1")

count = 0

#対象のシート
ids = [
"1kpfonHZjEsYNkn_C6x5rupoQGYbvbP5jlw1cOq6H44o",
"1KGtZpEhfvh2_xueldXoIQjmXrM_4AhLDT8G3PLWSSso",
]

ids.each{|id|

  ws = session.spreadsheet_by_key(id).worksheet_by_title("シート1")

  #今現在の最終行を取得
  for n in 3..200
    b = ws["A#{n}"]
    p b
    if b != "" then
    	rws[n + count,1] = ws["A#{n}"]
    	rws[n + count,2] = ws["B#{n}"]
    	rws[n + count,3] = ws["C#{n}"]
    	rws[n + count,4] = ws["D#{n}"]
    	rws[n + count,5] = ws["E#{n}"]
    	rws[n + count,6] = ws["F#{n}"]
    	rws[n + count,7] = ws["G#{n}"]
    	rws[n + count,8] = ws["H#{n}"]
    	rws[n + count,9] = ws["I#{n}"]
    	rws[n + count,10] = ws["J#{n}"]
    	rws[n + count,11] = ws["K#{n}"]
    	rws[n + count,12] = ws["L#{n}"]
    	rws[n + count,13] = ws["M#{n}"]
    	rws[n + count,14] = ws["N#{n}"]
    	rws[n + count,15] = ws["O#{n}"]
    	rws[n + count,16] = ws["P#{n}"]
    	rws[n + count,17] = ws["Q#{n}"]
    	rws[n + count,18] = ws["R#{n}"]
    	rws[n + count,19] = ws["S#{n}"]
    	rws[n + count,20] = ws["T#{n}"]
    	rws[n + count,21] = ws["U#{n}"]
    	rws[n + count,22] = ws["V#{n}"]
    	rws[n + count,23] = ws["W#{n}"]
    	rws[n + count,24] = ws["X#{n}"]
    	rws[n + count,25] = ws["Y#{n}"]
    	rws[n + count,26] = ws["Z#{n}"]
    	rws[n + count,27] = ws["AA#{n}"]
    	rws[n + count,28] = ws["AB#{n}"]
    	rws[n + count,29] = ws["AC#{n}"]
    	rws[n + count,30] = ws["AD#{n}"]
    	rws[n + count,31] = ws["AE#{n}"]
    	rws[n + count,32] = ws["AF#{n}"]
    	rws[n + count,33] = ws["AG#{n}"]
    	rws[n + count,34] = ws["AH#{n}"]
    	rws[n + count,35] = ws["AI#{n}"]
    	rws[n + count,36] = ws["AJ#{n}"]
    	rws[n + count,37] = ws["AK#{n}"]
    	rws[n + count,38] = ws["AL#{n}"]
    	rws[n + count,39] = ws["AM#{n}"]
    	rws[n + count,40] = ws["AN#{n}"]
    	rws[n + count,41] = ws["AO#{n}"]
    	rws[n + count,42] = ws["AP#{n}"]
    	rws[n + count,43] = ws["AQ#{n}"]
    	rws[n + count,44] = ws["AR#{n}"]
    	rws[n + count,45] = ws["AS#{n}"]
    	rws[n + count,46] = ws["AT#{n}"]
    	rws[n + count,47] = ws["AU#{n}"]
    	rws[n + count,48] = ws["AV#{n}"]
    	rws[n + count,49] = ws["AW#{n}"]
    	rws[n + count,50] = ws["AX#{n}"]
    	rws[n + count,51] = ws["AY#{n}"]
    	rws[n + count,52] = ws["AZ#{n}"]
    	rws[n + count,53] = ws["BA#{n}"]
    	rws[n + count,54] = ws["BB#{n}"]
    	rws[n + count,55] = ws["BC#{n}"]
    else
      count = n + count - 3
      break
    end

  end
  rws.save
}
