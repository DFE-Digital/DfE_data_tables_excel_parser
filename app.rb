# frozen_string_literal: true
require 'roo'

data_tables_path = './data/NPD Tables Oct18 20181010.xlsx'
data_tables_workbook = Roo::Spreadsheet.open(data_tables_path)

sheets_to_process = [
  {
    name: 'SC_Pupil_01-02_to_17-18_SUM', 
    data_blocks: [
      { header_row: 5, last_row: 94 },
      { header_row: 100, last_row: 137 }
    ]
  },{
    name: 'SC_Addresses_05-06_to_17-18_SUM', 
    data_blocks: [
      { header_row: 8, last_row: 23 }
    ]
  },{
    name: 'PRU_Census_09-10_to_12-13', 
    data_blocks: [
      { header_row: 6, last_row: 51 },
      { header_row: 59, last_row: 73 }
    ]
  },{
    name: 'EarlyYearsCensus_07-08_to_17-18', 
    data_blocks: [
      { header_row: 5, last_row: 60 },
      { header_row: 67, last_row: 135 }
    ]
  },{
    name: 'Alt_Provision_07-08_to_17-18', 
    data_blocks: [
      { header_row: 5, last_row: 50 },
      { header_row: 58, last_row: 67 }
    ]
  },{
    name: 'AP addresses_17_18', 
    data_blocks: [
      { header_row: 8, last_row: 23 }
    ]
  },{
    name: 'EYFSP_02-03_to_17-18', 
    data_blocks: [
      { header_row: 5, last_row: 26 },
      { header_row: 34, last_row: 193 }
    ]
  },{
    name: 'Phonics_11-12_to_17-18', 
    data_blocks: [
      { header_row: 5, last_row: 16 },
      { header_row: 24, last_row: 32 }
    ]
  },{
    name: 'KS1_97-98_to_17-18', 
    data_blocks: [
      { header_row: 5, last_row: 20 },
      { header_row: 28, last_row: 140 }
    ]
  },{
    name: 'KS2_95-96_to_17-18', 
    data_blocks: [
      { header_row: 5, first_row: 7, last_row: 91 },
      { header_row: 5, first_row: 94, last_row: 103 },
      { header_row: 110, first_row: 112, last_row: 507 },
      { header_row: 110, first_row: 510,last_row: 555 }
    ]
  },{
    name: 'Year_7_00-01_to_06-07', 
    data_blocks: [
      { header_row: 5, last_row: 17 },
      { header_row: 25, last_row: 59 }
    ]
  },{
    name: 'KS3_97-98_to_12-13', 
    data_blocks: [
      { header_row: 5, first_row: 7, last_row: 16, table_name: 'KS3_Candidate' },
      { header_row: 5, first_row: 28, last_row: 69, table_name: 'KS3_Indicators' },
      { header_row: 5, first_row: 72, last_row: 74, table_name: 'KS3_Results' },
      { header_row: 81, first_row: 83, last_row: 96, table_name: 'KS3_Candidate' },
      { header_row: 81, first_row: 99, last_row: 247, table_name: 'KS3_Indicators' },
      { header_row: 81, first_row: 250, last_row: 296, table_name: 'KS3_Results' },
    ]
  },{
    name: 'KS4_01-02_17-18', 
    data_blocks: [
      { header_row: 5, first_row: 7, last_row: 101, table_name: 'KS4_Pupil' },
      { header_row: 5, first_row: 104, last_row: 114, table_name: 'KS4_Exam' },
      { header_row: 121, first_row: 123, last_row: 1189, table_name: 'KS4_Pupil' },
      { header_row: 121, first_row: 1192, last_row: 1243, table_name: 'KS4_Exam' },
    ]
  },{
    name: 'KS5_01-02_to_17-18', 
    data_blocks: [
      { header_row: 5, first_row: 7, last_row: 42, table_name: 'KS5_Student' },
      { header_row: 5, first_row: 44, last_row: 53, table_name: 'KS5_Exam' },
      { header_row: 60, first_row: 62, last_row: 689, table_name: 'KS5_Student' },
      { header_row: 60, first_row: 692, last_row: 757, table_name: 'KS5_Exam' },
    ]
  },{
    name: 'CIN_08-09_to_16-17', 
    data_blocks: [
      { header_row: 8, first_row: 10, last_row: 10, table_name: 'n_Census_CIN_Overall' },
      { header_row: 8, first_row: 12, last_row: 50, table_name: 'n_Census_CIN_Child' },
      { header_row: 8, first_row: 52, last_row: 64, table_name: 'n_Census_CIN_CPP' },
      { header_row: 8, first_row: 66, last_row: 86, table_name: 'n_Census_CIN_Details' },
      { header_row: 8, first_row: 88, last_row: 92, table_name: 'n_Census_CIN_Disabilities' },
      { header_row: 8, first_row: 94, last_row: 98, table_name: 'n_Census_CIN_OpenCase' },
      { header_row: 8, first_row: 100, last_row: 108, table_name: 'n_Census_CIN_ServiceProvision' },
    ]
  },{
    name: 'CLA_05-06_to_16-17', 
    data_blocks: [
      { header_row: 8, last_row: 36 }
    ]
  },{
    name: 'Absence_05-06_to_17-18', 
    data_blocks: [
      { header_row: 5, last_row: 152 },
      { header_row: 159, last_row: 226 }
    ]
  },{
    name: 'Exclusions_01-02_to_04-05', 
    data_blocks: [
      { header_row: 5, last_row: 20 },
      { header_row: 27, last_row: 32 },
    ]
  },{
    name: 'Exclusions_05-06_to_16-17', 
    data_blocks: [
      { header_row: 5, last_row: 59 },
      { header_row: 5, first_row: 62, last_row: 68 },
      { header_row: 76, last_row: 96 },
      { header_row: 76, first_row: 99, last_row: 103 },
    ]
  },{
    name: 'PLAMS_07-08_to_16-17', 
    data_blocks: [
      { header_row: 5, last_row: 17 },
      { header_row: 24, last_row: 76 }
    ]
  },{
    name: 'NCCIS_10-11_to_16-17', 
    data_blocks: [
      { header_row: 5, last_row: 27 },
      { header_row: 35, last_row: 58 }
    ]
  },{
    name: 'ISP_09-10_to_12-13', 
    data_blocks: [
      { header_row: 5, first_row: 7, last_row: 18, table_name: 'TODO-ISP-Placement/Student' },
      { header_row: 25, first_row: 27, last_row: 62, table_name: 'TODO-ISP-Placement/Student' },
      { header_row: 25, first_row: 65, last_row: 79, table_name: 'TODO-ISP-Funding' },
      { header_row: 25, first_row: 82, last_row: 106, table_name: 'TODO-ISP-Support' },
    ]
  },{
    name: 'YPMAD 16-17', 
    data_blocks: [
      { header_row: 5, first_row: 7, last_row: 33, table_name: 'TODO-YPMAD-Chronological Indicators' },
      { header_row: 5, first_row: 36, last_row: 101, table_name: 'TODO-YPMAD-Snapshot Indicators' },
      { header_row: 109, first_row: 111, last_row: 274, table_name: 'TODO-YPMAD-Chronological Indicators' },
      { header_row: 109, first_row: 277, last_row: 375, table_name: 'TODO-YPMAD-Snapshot Indicators' },
    ]
  }
]


require 'elasticsearch'

client = Elasticsearch::Client.new url: 'http://localhost:9200', log: true

# client.transport.reload_connections!
# client.cluster.health

puts "Results:"
puts client.search q: 'test'

# For each worksheet
sheets_to_process.each do |sheet|
  puts sheet[:name]

  # For each block within a sheet
  sheet[:data_blocks].each do |block|
    block_name = block[:table_name] || sheet[:name]
    headers = data_tables_workbook.sheet(sheet[:name]).row(block[:header_row])

    # Cast empty strings to nil
    headers.map{|cell| (cell.instance_of?(String) && cell.empty?) ? nil : cell}

    # Remove nils from the end of the header list
    headers = headers.reverse.drop_while(&:nil?).reverse

    # The first row is either explicitly specified OR the row after the header
    first_row_number = block[:first_row] || (block[:header_row] +1)
    last_row_number = block[:last_row]

    
    (first_row_number..last_row_number).each do |row_number|
      row = data_tables_workbook.sheet(sheet[:name]).row(row_number)

      next if row.compact.empty?

      data_element = {
        group_name: sheet[:name],
        table_name: block_name,
      }

      row.each_with_index do |cell, index|
        # Don't collect (generally empty) cells outside the table
        next if headers[index].nil?

        # Cast empty strings to nil, so we don't break the ElasticSearch ingestion
        data_element[headers[index]] = (cell.instance_of?(String) && cell.empty?) ? nil : cell
      end

      # A really specific instance of merged cells in the CLA table on row 28
      next if sheet[:name] == 'CLA_05-06_to_16-17' && data_element["NPDAlias"].nil?

      # Post-process to add structure
      # TODO: cope with SUM SC Addresses
      data_element["Collection term"] = data_element["Collection term"]&.split(', ')

      npd_alias = data_element.delete("NPDAlias") || data_element.delete('NPD Alias ') || data_element["NPD Alias"]
      npd_alias = npd_alias.split("\n")
      data_element["NPD Alias"] = npd_alias

      

      # Years can be written as "2006/07 only. Coerce that into "2006/07 - 2006-07" for consistency.
      years_populated = data_element.delete("Years populated") || data_element["Years Populated"]

      # KS3 Result Table contains no years collected data
      if years_populated
        years_populated = years_populated.gsub(/(.*) only/, '\1 - \1')
        data_element["Years Populated"] = years_populated

        # Break up years into array [start, end]
        years_populated = years_populated.split(/ *- */)

        # Create collected_from for collection from Years Populated
        start_year = years_populated.first[0..3].to_i
        start_date = Date.new(start_year,9,1)
        data_element[:collected_from] = years_populated.first
        data_element[:collected_from_date] = start_date

        # Create collected_until for collection from Years Populated
        if years_populated.size == 2
          end_year = years_populated.last[0..3].to_i
          end_date = Date.new(end_year + 1 ,9, 1)
          data_element[:collected_until] = years_populated.last
          data_element[:collected_until_date] = end_date
        else
          data_element[:collected_until] = 'Present'
        end

        # Finally, coerce these into a full list of years
        # data_element["Years Populated"] = years_populated
      end
      
      client.index  index: 'data_elements', type: 'data_element', body: data_element
    end

    # puts "\t#{block_name}: #{data_tables_workbook.sheet(sheet[:name]).row(block[:header_row])[0]}"
  end
  # 

  #puts sheet.row(5).cell(1)
  #puts "---------"
end
