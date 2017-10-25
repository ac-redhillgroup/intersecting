require 'roo'
require 'json'
#create dictionaraies/hashes of the objects
fast_routes = %w[1 2 3 4 5 6 7 8 9 20 30 40 90 OTHER]
rvdb_routes = %w[50 52 OTHER]
st_routes = %w[1 2 3 4 5 6 7 8 9 15 17 20 78 80 82 85 OTHER]
vc_routes = %w[1 2 4 5 6 8 OTHER]
sagencies = %w[FS RV ST VC]
agencies = %w[3D AC AY BA CC FS GG RV SF SM ST VC VN WC WH OTHER]
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
		rte = "RV-" + rvdb_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	when "ST"
		rte = "ST-" + st_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	when "VC"
		rte = "VC-" + vc_routes_hash[sheet.row(line)[sroutes[agy.to_s]-1]]
	end
	id = sheet.row(line)[0].to_s
	#tranfer before
	if sheet.row(line)[29] == 1
		first_bf = agenices_hash[sheet.row(line)[30]]
		#first route before and current route
		if sheet.row(line)[31].nil? #if not other
			first_rte = first_bf.to_s + '-' + sheet.row(line)[32].to_s
			#BART LOGIC
			if first_bf == "BA"
				begin 
					unless json[first_bf].include?(rte)
						p "#{id} #{rte} | #{first_rte} | First & Current with Bart"
					end
 				rescue
 			  	"Error #{id}"
 				end
			else
				begin 
					unless json[first_rte].include?(rte)
 						p "#{id} : #{rte} | #{first_rte} First & Current w/o Bart"
 				  end
 				rescue
 			  	"Error #{id} #{rte}"
 				end
			end
		else
			first_rte = first_bf.to_s + '-' + sheet.row(line)[31].to_s
			#other logic..
		end
		#first route before and current end
		#second route before and first route
		if sheet.row(line)[36] == 1
 			second_bf = agenices_hash[sheet.row(line)[37]]
 			if sheet.row(line)[38].nil?
 				second_rte = second_bf.to_s + '-' + sheet.row(line)[39].to_s
 				if second_bf == 'BA' 
	 				begin 
						unless json[second_bf].include?(first_rte)
							p "#{id} #{first_rte} | #{second_bf} | Second & First with Bart"
						end
	 				rescue
	 			  	"Error #{id}"
	 				end
 				else
 					begin 
						unless json[second_rte].include?(second_rte)
							p "#{id} #{first_rte} | #{second_rte} | Second & First w/o Bart"
						end
 					rescue
 			  		"Error #{id}"
 					end
 				end
 			else
 				second_rte = second_bf.to_s + '-' + sheet.row(line)[38].to_s
 				#other logic
 			end
 			if sheet.row(line)[43] == 1
 				third_bf = agenices_hash[sheet.row(line)[44]]
 				#p "#{second_rte} | #{third_rte} #{json[third_rte].include?(second_rte)}"
 				if sheet.row(line)[45].nil?
 					third_rte = third_bf.to_s + '-' + sheet.row(line)[46].to_s
 					if third_bf == "BA"
 						begin 
							unless json[third_bf].include?(second_rte)
								p "#{id} #{second_rte} | #{third_rte} | Third & Second with Bart"
							end
	 					rescue
	 			  		"Error #{id}"
	 					end
 					else
 						begin 
							unless json[second_bf].include?(first_rte)
								p "#{id} #{second_rte} | #{third_rte} | Third & Second with Bart"
							end
	 					rescue
	 			  		"Error #{id}"
	 					end
 					end
 				else
 					#other
 				end
			end
		end
  end


	if sheet.row(line)[55] == 1
		first_af = agenices_hash[sheet.row(line)[56]]
		#first route before and current route
		if sheet.row(line)[57].nil? #if not other
			first_rte = first_af.to_s + '-' + sheet.row(line)[58].to_s
			#BART LOGIC
			if first_af == "BA"
				begin 
					unless json[first_af].include?(rte)
						p "#{id} #{rte} | #{first_rte} | First & Current with Bart"
					end
 				rescue
 			  	"Error #{id}"
 				end
			else
				begin 
					unless json[first_rte].include?(rte)
 						p "#{id} : #{rte} | #{first_rte} First & Current w/o Bart After"
 				  end
 				rescue
 			  	"Error #{id} #{rte}"
 				end
			end
		else
			first_rte = first_af.to_s + '-' + sheet.row(line)[57].to_s
			#other logic..
		end
		#first route after and current end
		#second route after and first route
		if sheet.row(line)[62] == 1
 			second_af = agenices_hash[sheet.row(line)[63]]
 			if sheet.row(line)[64].nil?
 				second_rte = second_af.to_s + '-' + sheet.row(line)[65].to_s
 				if second_af == 'BA' 
	 				begin 
						unless json[second_af].include?(first_rte)
							p "#{id} #{first_rte} | #{second_af} | Second & First with Bart"
						end
	 				rescue
	 			  	"Error #{id}"
	 				end
 				else
 					begin 
						unless json[second_rte].include?(second_rte)
							p "#{id} #{first_rte} | #{second_rte} | Second & First w/o Bart"
						end
 					rescue
 			  		"Error #{id}"
 					end
 				end
 			else
 				second_rte = second_af.to_s + '-' + sheet.row(line)[64].to_s
 				#other logic
 			end
 			if sheet.row(line)[69] == 1
 				third_af = agenices_hash[sheet.row(line)[70]]
 				#p "#{second_rte} | #{third_rte} #{json[third_rte].include?(second_rte)}"
 				if sheet.row(line)[71].nil?
 					third_rte = third_af.to_s + '-' + sheet.row(line)[72].to_s
 					if third_af == "BA"
 						begin 
							unless json[third_af].include?(second_rte)
								p "#{id} #{second_rte} | #{third_rte} | Third & Second with Bart"
							end
	 					rescue
	 			  		"Error #{id}"
	 					end
 					else
 						begin 
							unless json[second_af].include?(first_rte)
								p "#{id} #{second_rte} | #{third_rte} | Third & Second with Bart"
							end
	 					rescue
	 			  		"Error #{id}"
	 					end
 					end
 				else
 					#other
 				end
			end
		end
  end  



	# # first_af = agenices_hash[sheet.row(line)[56]]
 #    # print first_bf
end