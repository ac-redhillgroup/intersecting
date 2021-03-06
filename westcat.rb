require 'roo'
require 'csv'
require 'json'

THIRD_TRANSIT_BEFORE_USED_OR_NOT = 41
THIRD_TRANSIT_BEFORE = 42
THIRD_TRANSIT_OTHER = 43
THIRD_TRANSIT_ROUTE = 44

SECOND_TRANSIT_BEFORE_USED_OR_NOT = 33
SECOND_TRANSIT_BEFORE = 34
SECOND_TRANSIT_OTHER = 35
SECOND_TRANSIT_ROUTE = 36

FIRST_TRANSIT_BEFORE_USED_OR_NOT = 25
FIRST_TRANSIT_BEFORE = 26
FIRST_TRANSIT_OTHER = 27
FIRST_TRANSIT_ROUTE = 28

FIRST_TRANSIT_AFTER_USED_OR_NOT = 56
FIRST_TRANSIT_AFTER = 57
FIRST_TRANSIT_OTHER_AFTER = 58
FIRST_TRANSIT_ROUTE_AFTER = 59

SECOND_TRANSIT_AFTER_USED_OR_NOT = 64
SECOND_TRANSIT_AFTER = 65
SECOND_TRANSIT_OTHER_AFTER = 66
SECOND_TRANSIT_ROUTE_AFTER = 67

THIRD_TRANSIT_AFTER_USED_OR_NOT = 72
THIRD_TRANSIT_AFTER = 73
THIRD_TRANSIT_OTHER_AFTER = 74
THIRD_TRANSIT_ROUTE_AFTER = 75

class ExportData

	def initialize
		#reading intersecting dict and parsing it
		@json = JSON.parse(File.read("intersect_dict_json.txt"))
		#create dictionaraies/hashes of the objects
		@errors = []

		@sagencies = %w[10 11 12 15 16 17 18 19 30Z C3 JR JL JX JPX LYNX OTHER]
		@agencies = %w[3D AC EM BA CC FS GG SF SM ST VN WC OTHER]
		
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
		#reading spreadsheet
		workbook = Roo::Spreadsheet.open("MTCWESTCAT.xlsx")
		sheet = workbook.sheet(0)
		(2..sheet.last_row).each do |line|
		 	#this piece of code is to get the current agency and route
			agy = sheet.row(line)[2]
			curr = @sagencies_hash[agy]
			rte = "WC-" + curr
			id = sheet.row(line)[0].to_s
			#tranfer before
			if sheet.row(line)[THIRD_TRANSIT_BEFORE_USED_OR_NOT] == 1
				third_rte = find_record(THIRD_TRANSIT_BEFORE,THIRD_TRANSIT_OTHER,THIRD_TRANSIT_ROUTE,sheet,rte,line,id)
				second_rte = find_record(SECOND_TRANSIT_BEFORE,SECOND_TRANSIT_OTHER,SECOND_TRANSIT_ROUTE,sheet,third_rte,line,id)
				first_rte = find_record(FIRST_TRANSIT_BEFORE,FIRST_TRANSIT_OTHER,FIRST_TRANSIT_ROUTE,sheet,second_rte,line,id)
			elsif sheet.row(line)[SECOND_TRANSIT_BEFORE_USED_OR_NOT] == 1
				second_rte = find_record(SECOND_TRANSIT_BEFORE,SECOND_TRANSIT_OTHER,SECOND_TRANSIT_ROUTE,sheet,rte,line,id)
				first_rte = find_record(FIRST_TRANSIT_BEFORE,FIRST_TRANSIT_OTHER,FIRST_TRANSIT_ROUTE,sheet,second_rte,line,id)
			elsif sheet.row(line)[FIRST_TRANSIT_BEFORE_USED_OR_NOT] == 1
				first_rte = find_record(FIRST_TRANSIT_BEFORE,FIRST_TRANSIT_OTHER,FIRST_TRANSIT_ROUTE,sheet,rte,line,id)
			end

			#transfer after
			if sheet.row(line)[FIRST_TRANSIT_AFTER_USED_OR_NOT] == 1
				first_rte = find_record(FIRST_TRANSIT_AFTER,FIRST_TRANSIT_OTHER_AFTER,FIRST_TRANSIT_ROUTE_AFTER,sheet,rte,line,id)
				if sheet.row(line)[SECOND_TRANSIT_AFTER_USED_OR_NOT] == 1
		 			second_rte = find_record(SECOND_TRANSIT_AFTER,SECOND_TRANSIT_OTHER_AFTER,SECOND_TRANSIT_ROUTE_AFTER,sheet,first_rte,line,id)
		 			if sheet.row(line)[THIRD_TRANSIT_AFTER_USED_OR_NOT] == 1
		 				third_rte = find_record(THIRD_TRANSIT_AFTER,THIRD_TRANSIT_OTHER_AFTER,THIRD_TRANSIT_ROUTE_AFTER,sheet,second_rte,line,id)
		 			end
				end
		  end  
		 end
		 generate_csv()
	end

	def find_record(tty,tto,ttr,sheet,rte,line,id)
		# p line
		tranfer_type = @agenices_hash[sheet.row(line)[tty]]  #code might change for WS
		
			if sheet.row(line)[tto].nil?
				if tranfer_type == "BA"
					transfer_rte = "BA"
					evaluate(rte,transfer_rte,id)
				else
					transfer_rte = tranfer_type.to_s + '-' + sheet.row(line)[ttr].to_s
					evaluate(rte,transfer_rte,id)
				end
			else
				transfer_rte = tranfer_type.to_s + '-' + sheet.row(line)[tto].to_s
				evaluate(rte,transfer_rte,id)
			end
			return transfer_rte
	end

private
	def evaluate(rte,transfer_rte,id)
		begin 
		  unless @json[transfer_rte].include?(rte)
				p "#{id} #{rte} | #{transfer_rte} "
				@errors << id
				@errors << rte 
				@errors << transfer_rte
		  end
	 	rescue
		  p "Error #{id} | #{rte} | #{transfer_rte}"
		  @errors << id
		  @errors << rte 
		  @errors << transfer_rte
	  	end
	end

	def generate_csv
	 #change the number according to the data exported
	 @errs = @errors.each_slice(3)
	 CSV.open('data.csv','wb') do |csv|
	 	 csv << ["ID", "Route", "Route2"]
	   @errs.each do |error|
	     csv << error
	   end
	 end
	end
end
ExportData.new.export_data()