require 'roo'
require 'json'
#create dictionaraies/hashes of the objects
fast_routes = %w[1 2 3 4 5 6 7 8 9 20 30 40 90 OTHER]
rvdb_routes = %w[50 52 OTHER]
st_routes = %w[1 2 3 4 5 6 7 8 9 15 17 20 78 80 82 85 OTHER]
vc_routes = %w[1 2 4 5 6 8 OTHER]
sagencies = %w[FS RV ST VC]
agencies = %w[3D AC VN BA CC FS GG RV SF SM ST VC AY WC WH OTHER]
sroutes = {"1" => 4 , "2" => 6,"3" => 8, "4" => 10}

rvdb_routes_hash = {}
rvdb_routes.each.with_index(1) do |route,idx|
  rvdb_routes_hash[idx] = route
end

fast_routes_hash = {}
fast_routes.each.with_index(1) do |route,idx|
  fast_routes_hash[idx] = route
end

st_routes_hash = {}
st_routes.each.with_index(1) do |route,idx|
  st_routes_hash[idx] = route
end

vc_routes_hash = {}
vc_routes.each.with_index(1) do |route,idx|
  vc_routes_hash[idx] = route
end

agenices_hash = {}
agencies.each.with_index(1) do |route,idx|
  agenices_hash[idx] = route
end

sagencies_hash = {}
sagencies.each.with_index(1) do |route,idx|
  sagencies_hash[idx] = route
end

# [sagencies_hash,fast_routes_hash,vc_routes_hash,st_routes_hash,rvdb_routes_hash].each do |dic|
# 	print dic
# 	puts ""
# end

f = File.read("intersect_dict_json.txt")
json = JSON.parse(f)

workbook = Roo::Spreadsheet.open("SOLANO.xlsx")
sheet = workbook.sheet(0)
(2..sheet.last_row).each do |line|
 	#this piece of code is to get the current agency and route
	agy = sheet.row(line)[2]
	curr = sagencies_hash[agy]
	case curr
	when "FS"
		rte = "FS-" + fast_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	when "RV"
		rte = "RV`-" + rvdb_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	when "ST"
		rte = "ST-" + st_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	when "VC"
		rte = "VC-" + vc_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	end
	
	#tranfer before
	if sheet.row(line)[29] == 1
		first_bf = agenices_hash[sheet.row(line)[30]]
		if sheet.row(line)[36] == 1
 			second_bf = agenices_hash[sheet.row(line)[37]]
 				if sheet.row(line)[43] == 1
				    third_bf = agenices_hash[sheet.row(line)[44]]
				end
		end
    end

    #tranfer after
    if sheet.row(line)[55] == 1
		first_bf = agenices_hash[sheet.row(line)[56]]
		if sheet.row(line)[62] == 1
 			second_bf = agenices_hash[sheet.row(line)[63]]
 				if sheet.row(line)[69] == 1
				    third_bf = agenices_hash[sheet.row(line)[70]]
				end
		end
    end

	# first_af = agenices_hash[sheet.row(line)[56]]
    # print first_bf
end