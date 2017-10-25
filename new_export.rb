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

class ExportData

	def initialize
		#create dictionaraies/hashes of the objects
		@fast_routes = %w[1 2 3 4 5 6 7 8 9 20 30 40 90 OTHER]
		@rvdb_routes = %w[50 52 OTHER]
		@st_routes = %w[1 2 3 4 5 6 7 8 9 15 17 20 78 80 82 85 OTHER]
		@vc_routes = %w[1 2 4 5 6 8 OTHER]
		@sagencies = %w[FS RV ST VC]
		@agencies = %w[3D AC AY BA CC FS GG RV SF SM ST VC VN WC WH OTHER]
		@sroutes = {"1" => 4 , "2" => 6,"3" => 8, "4" => 10}

		@rvdb_routes_hash = {}
		@rvdb_routes.each.with_index(1) do |route,idx|
		  @rvdb_routes_hash[idx] = route
		end

		@fast_routes_hash = {}
		@fast_routes.each.with_index(1) do |route,idx|
		  @fast_routes_hash[idx] = route
		end

		@st_routes_hash = {}
		@st_routes.each.with_index(1) do |route,idx|
		  @st_routes_hash[idx] = route
		end

		@vc_routes_hash = {}
		@vc_routes.each.with_index(1) do |route,idx|
		  @vc_routes_hash[idx] = route
		end

		@agenices_hash = {}
		@agencies.each.with_index(1) do |route,idx|
		  @agenices_hash[idx] = route
		end

		@sagencies_hash = {}
		@sagencies.each.with_index(1) do |route,idx|
		  @sagencies_hash[idx] = route
		end
	end
	
	def export_data
		f = File.read("intersect_dict_json.txt")
		json = JSON.parse(f)
		workbook = Roo::Spreadsheet.open("SOLANO.xlsx")
		sheet = workbook.sheet(0)
		(2..sheet.last_row).each do |line|
		 	#this piece of code is to get the current agency and route
			agy = sheet.row(line)[2]
			curr = @sagencies_hash[agy]
			case curr
			when "FS"
				rte = "FS-" + @fast_routes_hash[sheet.row(line)[@sroutes[agy.to_s]-1]]
			when "RV"
				rte = "RV-" + @rvdb_routes_hash[sheet.row(line)[@sroutes[agy.to_s]-1]]
			when "ST"
				rte = "ST-" + @st_routes_hash[sheet.row(line)[@sroutes[agy.to_s]-1]]
			when "VC"
				rte = "VC-" + @vc_routes_hash[sheet.row(line)[@sroutes[agy.to_s]-1]]
			end
			id = sheet.row(line)[0].to_s
			#tranfer before
			if sheet.row(line)[THIRD_TRANSIT_BEFORE_USED_OR_NOT] == 1
				third_rte = find_record(THIRD_TRANSIT_BEFORE,THIRD_TRANSIT_OTHER,THIRD_TRANSIT_ROUTE,sheet,rte,line,id,json)
				second_rte = find_record(SECOND_TRANSIT_BEFORE,SECOND_TRANSIT_OTHER,SECOND_TRANSIT_ROUTE,sheet,third_rte,line,id,json)
				first_rte = find_record(FIRST_TRANSIT_BEFORE,FIRST_TRANSIT_OTHER,FIRST_TRANSIT_ROUTE,sheet,second_rte,line,id,json)
			elsif sheet.row(line)[SECOND_TRANSIT_BEFORE_USED_OR_NOT] == 1
				second_rte = find_record(SECOND_TRANSIT_BEFORE,SECOND_TRANSIT_OTHER,SECOND_TRANSIT_ROUTE,sheet,rte,line,id,json)
				first_rte = find_record(FIRST_TRANSIT_BEFORE,FIRST_TRANSIT_OTHER,FIRST_TRANSIT_ROUTE,sheet,second_rte,line,id,json)
			elsif sheet.row(line)[FIRST_TRANSIT_BEFORE_USED_OR_NOT] == 1
				first_rte = find_record(FIRST_TRANSIT_BEFORE,FIRST_TRANSIT_OTHER,FIRST_TRANSIT_ROUTE,sheet,rte,line,id,json)
			end

			if sheet.row(line)[FIRST_TRANSIT_AFTER_USED_OR_NOT] == 1
				first_rte = find_record(FIRST_TRANSIT_AFTER,FIRST_TRANSIT_OTHER_AFTER,FIRST_TRANSIT_ROUTE_AFTER,sheet,rte,line,id,json)
				#second route after and first route
				if sheet.row(line)[SECOND_TRANSIT_AFTER_USED_OR_NOT] == 1
		 			second_rte = find_record(SECOND_TRANSIT_AFTER,SECOND_TRANSIT_OTHER_AFTER,SECOND_TRANSIT_ROUTE_AFTER,sheet,first_rte,line,id,json)
		 			if sheet.row(line)[THIRD_TRANSIT_AFTER_USED_OR_NOT] == 1
		 				third_rte = find_record(THIRD_TRANSIT_AFTER,THIRD_TRANSIT_OTHER_AFTER,THIRD_TRANSIT_ROUTE_AFTER,sheet,second_rte,line,id,json)
		 			end
				end
		  end  
		 end
	end

def find_record(tty,tto,ttr,sheet,rte,line,id,json)
	tranfer_type = @agenices_hash[sheet.row(line)[tty]]
		
		if sheet.row(line)[tto].nil?
			if tranfer_type == "BA"
				transfer_rte = "BA"
				evaluate(json,rte,transfer_rte,id)
			else
				transfer_rte = tranfer_type.to_s + '-' + sheet.row(line)[ttr].to_s
				evaluate(json,rte,transfer_rte,id)
			end
		else
			transfer_rte = tranfer_type.to_s + '-' + sheet.row(line)[tto].to_s
			evaluate(json,rte,transfer_rte,id)
		end
		return transfer_rte
end

def evaluate(json,rte,transfer_rte,id)
	begin 
	  unless json[transfer_rte].include?(rte)
		p "#{id} #{rte} | #{transfer_rte} "
	  end
  rescue
	  p "Error #{id} | #{rte} | #{transfer_rte}"
  end
end

end

ExportData.new.export_data()