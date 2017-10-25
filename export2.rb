require 'roo'
require 'csv'
require 'json'

THIRD_TRANSIT_BEFORE_USED_OR_NOT = 43
THIRD_TRANSIT_BEFORE = 44
THIRD_TRANSIT_OTHER = 45
THIRD_TRANSIT_ROUTE = 46

SECOND_TRANSIT_BEFORE_USED_OR_NOT = 36
SECOND_TRANSIT_BEFORE = 37
SECOND_TRANSIT_OTHER = 38
SECOND_TRANSIT_ROUTE = 39

FIRST_TRANSIT_BEFORE_USED_OR_NOT = 29
FIRST_TRANSIT_BEFORE = 30
FIRST_TRANSIT_OTHER = 31
FIRST_TRANSIT_ROUTE = 32

FIRST_TRANSIT_AFTER_USED_OR_NOT = 55
FIRST_TRANSIT_AFTER = 56
FIRST_TRANSIT_OTHER_AFTER = 57
FIRST_TRANSIT_ROUTE_AFTER = 58

SECOND_TRANSIT_AFTER_USED_OR_NOT = 62
SECOND_TRANSIT_AFTER = 63
SECOND_TRANSIT_OTHER_AFTER = 64
SECOND_TRANSIT_ROUTE_AFTER = 65

THIRD_TRANSIT_AFTER_USED_OR_NOT = 69
THIRD_TRANSIT_AFTER = 70
THIRD_TRANSIT_OTHER_AFTER = 71
THIRD_TRANSIT_ROUTE_AFTER = 72


def export_data
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
	if sheet.row(line)[THIRD_TRANSIT_BEFORE_USED_OR_NOT] == 1
		third_bf = agenices_hash[sheet.row(line)[THIRD_TRANSIT_BEFORE]]
		
		if sheet.row(line)[THIRD_TRANSIT_OTHER].nil?
			if third_bf == "BA"
				third_rte = "BA"
				begin 
					unless json[third_bf].include?(rte)
						p "#{id} #{rte} | #{third_rte} | Third & Current"
						# csv << [id,rte,third_rte]
					end
				rescue
	  			p "Error #{id} Before 3rd  #{rte} | #{third_rte}"
	  			  # csv << [id,rte,third_rte] 
				end
			else
				third_rte = third_bf.to_s + '-' + sheet.row(line)[THIRD_TRANSIT_ROUTE].to_s
				begin 
					unless json[third_rte].include?(rte)
						p "#{id} #{rte} | #{third_rte} | Third & Current"
						# csv << [id,rte,third_rte]
					end
				rescue
	  			p "Error #{id} Before 3rd  #{rte} | #{third_rte}"
	  				# csv << [id,rte,third_rte]
				end
			end
		else
			third_rte = third_bf.to_s + '-' + sheet.row(line)[THIRD_TRANSIT_OTHER].to_s
			begin 
				unless json[third_bf].include?(rte)
					p "#{id} #{rte} | #{third_rte} | Third & Current"
					# csv << [id,rte,third_rte]
				end
			rescue
	  			p "Error #{id} #{rte} | #{third_rte} | Third & Current OTHER"
	  			# csv << [id,rte,third_rte]
			end
		end

		#SECOND BEFORE
		second_bf = agenices_hash[sheet.row(line)[SECOND_TRANSIT_BEFORE]]
 			if sheet.row(line)[SECOND_TRANSIT_OTHER].nil?
 				if second_bf == 'BA' 
 					second_rte = 'BA'
	 				begin 
						unless json[second_bf].include?(third_rte)
							p "#{id} #{third_rte} | #{second_rte} | Third & Second"
							# csv << [id,second_rte,third_rte]
						end
	 				rescue
	 			  	 p "Error #{id} #{third_rte} | #{second_rte} | Third & Second"
	 			  	 # csv << [id,second_rte,third_rte]
	 				end
 				else
 					second_rte = second_bf.to_s + '-' + sheet.row(line)[SECOND_TRANSIT_ROUTE].to_s
 					begin 
						unless json[second_rte].include?(third_rte)
							p "#{id} #{third_rte} | #{second_rte} | Third & Second"
							# csv << [id,second_rte,third_rte]
						end
 					rescue
 			  		p "Error #{id} #{third_rte} | #{second_rte} | Third & Second"
 			  			# csv << [id,second_rte,third_rte]
 					end
 				end
 			else
 				second_rte = second_bf.to_s + '-' + sheet.row(line)[SECOND_TRANSIT_OTHER].to_s
	 			begin 
					unless json[second_rte].include?(third_rte)
						p "#{id} #{second_rte} | #{third_rte} | Third & Second"
						# csv << [id,second_rte,third_rte]
					end
				rescue
		  			p "Error #{id} #{second_rte} | #{third_rte} | Third & Second Before 2nd OTHER"
		  			# csv << [id,second_rte,third_rte]
				end
 			end

		#FIRST BEFORE
		first_bf = agenices_hash[sheet.row(line)[FIRST_TRANSIT_BEFORE]]
		#first route before and current route
		if sheet.row(line)[FIRST_TRANSIT_OTHER].nil? #if not other
			#BART LOGIC
			if first_bf == "BA"
				first_rte = "BA"
				begin 
					unless json[first_bf].include?(second_rte)
						p "#{id} #{first_rte} | #{second_rte} | First & Second"
						#csv << [id,first_rte,second_rte]
					end
 				rescue
 			  	p "Error #{id} #{id} #{first_rte} | #{second_rte} | First & Second Before 1st"
 			  	#csv << [id,first_rte,second_rte]
 				end
			else
				first_rte = first_bf.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_ROUTE].to_s
				begin 
					unless json[first_rte].include?(second_rte)
 						p "#{id} #{first_rte} | #{second_rte} | First & Second"
 						#csv << [id,first_rte,second_rte]
 				  end
 				rescue
 			  	p "Error #{id} #{first_rte} | #{second_rte} | First & Second Before 1st"
 			  	#csv << [id,first_rte,second_rte]
 				end
			end
		else
			first_rte = first_bf.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_OTHER].to_s
			begin 
				unless json[first_rte].include?(second_rte)
					p "#{id} #{first_rte} | #{second_rte} | First & Second"
					csv << [id,first_rte,second_rte]
				end
			rescue
	  			p "Error #{id} #{first_rte} | #{second_rte} | First & Second Before 2nd OTHER"
	  			csv << [id,first_rte,second_rte]
			end
		end

	elsif sheet.row(line)[SECOND_TRANSIT_BEFORE_USED_OR_NOT] == 1
			second_bf = agenices_hash[sheet.row(line)[SECOND_TRANSIT_BEFORE]]
 			if sheet.row(line)[SECOND_TRANSIT_OTHER].nil?
 				if second_bf == 'BA' 
 						second_rte = 'BA'
	 				begin 
						unless json[second_bf].include?(rte)
							p "#{id} #{rte} | #{second_rte} | Second & Current with Bart"
							csv << [id,rte,second_rte]
						end
	 				rescue
	 			  	p "Error #{id} #{rte} | #{second_rte} | Second & Current"
	 			  	csv << [id,rte,second_rte]
	 				end
 				else
 					second_rte = second_bf.to_s + '-' + sheet.row(line)[SECOND_TRANSIT_ROUTE].to_s
 					begin 
						unless json[second_rte].include?(rte)
							p "#{id} #{rte} | #{second_rte} | Second & Current w/o Bart"
							csv << [id,rte,second_rte]
						end
 					rescue
 			  		p "Error #{id} #{rte} | #{second_rte} | Second & Current"
 			  		csv << [id,rte,second_rte]
 					end
 				end
 			else
 				second_rte = second_bf.to_s + '-' + sheet.row(line)[SECOND_TRANSIT_OTHER].to_s
	 			begin 
					unless json[second_rte].include?(rte)
						p "#{id} #{second_rte} | #{rte} | Second & Current"
						csv << [id,rte,second_rte]
					end
				rescue
		  			p "Error #{id} #{second_rte} | #{rte} | Second & Current Before 2nd OTHER"
		  			csv << [id,rte,second_rte]
				end
 			end

			#FIRST BEFORE
			first_bf = agenices_hash[sheet.row(line)[FIRST_TRANSIT_BEFORE]]
			#first route before and current route
			if sheet.row(line)[FIRST_TRANSIT_OTHER].nil? #if not other
				#BART LOGIC
				if first_bf == "BA"
					first_rte = 'BA'
					begin 
						unless json[first_bf].include?(second_rte)
							p "#{id} #{second_rte} | #{first_rte} | Second & first with Bart"
							csv << [id,first_rte,second_rte]
						end
	 				rescue
	 			  	p "Error #{id} #{second_rte} | #{first_rte} | Second First with"
	 				end
				else
					first_rte = first_bf.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_ROUTE].to_s
					begin 
						unless json[first_rte].include?(second_rte)
	 						p "#{id} : #{second_rte} | #{first_rte}  Second & First w/o Bart"
	 						#csv << [id,first_rte,second_rte]
	 				  end
	 				rescue
	 			  	p "Error #{id} #{second_rte} | #{first_rte} "
	 			  	#csv << [id,first_rte,second_rte]
	 				end
				end
			else
				first_rte = first_bf.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_OTHER].to_s
				begin 
					unless json[first_rte].include?(second_rte)
						p "#{id} #{second_rte} | #{first_rte} | Second & First"
						#csv << [id,first_rte,second_rte]
					end
				rescue
		  			p "Error #{id} #{second_rte} | #{first_rte} | Second & First Before 1st OTHER"
		  			#csv << [id,first_rte,second_rte]
				end
			end
	elsif sheet.row(line)[FIRST_TRANSIT_BEFORE_USED_OR_NOT] == 1
		first_bf = agenices_hash[sheet.row(line)[FIRST_TRANSIT_BEFORE]]
		#first route before and current route
		if sheet.row(line)[FIRST_TRANSIT_OTHER].nil? #if not other
			
			#BART LOGIC
			if first_bf == "BA"
				first_rte = "BA"
				begin 
					unless json[first_bf].include?(rte)
						p "#{id} #{rte} | #{first_bf} | Current & First"
						#csv << [id,rte,first_rte]
					end
 				rescue
 			  	p "Error #{id}:  #{rte} | #{first_rte}  Current & First "
 			  	#csv << [id,rte,first_rte]
 				end
			else
				first_rte = first_bf.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_ROUTE].to_s
				begin 
					unless json[first_rte].include?(rte)
 						p "#{id} : #{rte} | #{first_rte}  Current & First "
 						#csv << [id,rte,first_rte]
 				  end
 				rescue
 			  	p "Error #{id}: #{rte} | #{first_rte}  Current & First "
 			  	#csv << [id,rte,first_rte]
 				end
			end
		else
			first_rte = first_bf.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_OTHER].to_s
			begin 
				unless json[first_rte].include?(rte)
					p "#{id} #{rte} | #{first_rte} | Current & First"
					#csv << [id,rte,first_rte]
				end
			rescue
	  			p "Error #{id} #{rte} | #{first_rte} | Current & First Before 1st OTHER"
	  			#csv << [id,rte,first_rte]
			end
		end
	end

	if sheet.row(line)[FIRST_TRANSIT_AFTER_USED_OR_NOT] == 1
		first_af = agenices_hash[sheet.row(line)[FIRST_TRANSIT_AFTER]]
		#first route before and current route
		if sheet.row(line)[FIRST_TRANSIT_OTHER_AFTER].nil? #if not other
			
			#BART LOGIC
			if first_af == "BA"
				first_rte = "BA"
				begin 
					unless json[first_af].include?(rte)
						p "#{id} #{rte} | #{first_rte} | Current & First"
						##csv << [id,rte,first_rte]
					end
 				rescue
 			  	p "Error #{id} #{rte} | #{first_rte} | Current & First"
 			  	##csv << [id,rte,first_rte]
 				end
			else
				first_rte = first_af.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_ROUTE_AFTER].to_s
				begin 
					unless json[first_rte].include?(rte)
 						p "#{id} : #{rte} | #{first_rte} Current & first"
 						##csv << [id,rte,first_rte]
 				  end
 				rescue
 			  	p "Error #{id}: #{rte} | #{first_rte} Current & first"
 			  	##csv << [id,rte,first_rte]
 				end
			end
		else
			first_rte = first_af.to_s + '-' + sheet.row(line)[FIRST_TRANSIT_OTHER_AFTER].to_s
			begin 
					unless json[first_rte].include?(rte)
 						p "#{id} : #{rte} | #{first_rte} Current & First"
 						##csv << [id,rte,first_rte]
 				  end
 				rescue
 			  	p "Error #{id} #{rte} | #{first_rte} Current & First : First After Other"
 			  	##csv << [id,rte,first_rte]
 				end
		end
		#first route after and current end
		#second route after and first route
		if sheet.row(line)[SECOND_TRANSIT_AFTER_USED_OR_NOT] == 1
 			second_af = agenices_hash[sheet.row(line)[SECOND_TRANSIT_AFTER]]
 			if sheet.row(line)[SECOND_TRANSIT_OTHER_AFTER].nil?
 				if second_af == 'BA' 
 					second_rte = "BA"
	 				begin 
						unless json[second_af].include?(first_rte)
							p "#{id} #{first_rte} | #{second_af} | First & Second "
							#csv << [id,first_rte,second_rte]
						end
	 				rescue
	 			  	p "Error #{id} #{first_rte} | #{second_af} | First & Second "
	 			  	#csv << [id,first_rte,second_rte]
	 				end
 				else
 					second_rte = second_af.to_s + '-' + sheet.row(line)[SECOND_TRANSIT_ROUTE_AFTER].to_s
 					begin 
						unless json[second_rte].include?(first_rte)
							p "#{id} #{first_rte} | #{second_rte} | First & Second"
							#csv << [id,first_rte,second_rte]
						end
 					rescue
 			  		p "Error #{id} #{first_rte} | #{second_rte} | First & Second : 2nd after"
 			  		#csv << [id,first_rte,second_rte]
 					end
 				end
 			else
 				second_rte = second_af.to_s + '-' + sheet.row(line)[SECOND_TRANSIT_OTHER_AFTER].to_s
 				begin 
						unless json[second_rte].include?(first_rte)
							p "#{id} #{first_rte} | #{second_rte} | First & Second"
							#csv << [id,first_rte,second_rte]
						end
 					rescue
 			  		p "Error #{id} #{first_rte} | #{second_rte} | First & Second : 2nd after OTHER"
 			  		#csv << [id,first_rte,second_rte]
 					end
 			end
 			if sheet.row(line)[THIRD_TRANSIT_AFTER_USED_OR_NOT] == 1
 				third_af = agenices_hash[sheet.row(line)[THIRD_TRANSIT_AFTER]]
 				#p "#{second_rte} | #{third_rte} #{json[third_rte].include?(second_rte)}"
 				if sheet.row(line)[THIRD_TRANSIT_OTHER_AFTER].nil?
 					
 					if third_af == "BA"
 						third_rte = "BA"
 						begin 
							unless json[third_af].include?(second_rte)
								p "#{id} #{second_rte} | #{third_af} | Second & third"
								#csv << [id,second_rte,third_rte]
							end
	 					rescue
	 			  		p "Error #{id} #{second_rte} | #{third_af} 3rd AFTER"
	 			  		#csv << [id,second_rte,third_rte]
	 					end
 					else
 						third_rte = third_af.to_s + '-' + sheet.row(line)[THIRD_TRANSIT_ROUTE_AFTER].to_s
 						begin 
							unless json[third_rte].include?(second_rte)
								p "#{id} #{second_rte} | #{third_rte} | Second & Third"
								#csv << [id,second_rte,third_rte]
							end
	 					rescue
	 			  		p "Error #{id} 3rd AFTER #{second_rte} | #{third_rte}"
	 			  		#csv << [id,second_rte,third_rte]
	 					end
 					end
 				else
 					third_rte = third_af.to_s + '-' + sheet.row(line)[THIRD_TRANSIT_OTHER_AFTER].to_s
 					begin 
						unless json[third_rte].include?(second_rte)
							p "#{id} #{second_rte} | #{third_rte} | Second & Third"
							#csv << [id,second_rte,third_rte]
						end
	 					rescue
	 			  		p "Error #{id} 3rd AFTER #{second_rte} | #{third_rte}"
	 			  		#csv << [id,second_rte,third_rte]
	 					end
 				end
			end
		end
  end  
 end
	
end

export_data