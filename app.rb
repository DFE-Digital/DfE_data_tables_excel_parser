# frozen_string_literal: true
require 'roo'

data_tables_path = '/Volumes/Twilight/Users/rb/Documents/Slate Horse/clients/Nimble/2018-11-17 NPD post-alpha/data_tables_parser/data/NPD Tables Oct18 20181010.xlsx'
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
      { header_row: 5, last_row: 103 },
      { header_row: 110, last_row: 555 }
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
      { header_row: 5, first_row: 27, last_row: 69, table_name: 'KS3_Indicators' },
      { header_row: 5, first_row: 71, last_row: 74, table_name: 'KS3_Results' },
      { header_row: 81, first_row: 83, last_row: 96, table_name: 'KS3_Candidate' },
      { header_row: 81, first_row: 98, last_row: 247, table_name: 'KS3_Indicators' },
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
      { header_row: 5, last_row: 68 },
      { header_row: 76, last_row: 103 },
    ]
  },{
    name: 'PLAMS_07-08_to_16-17', 
    data_blocks: [
      { header_row: 5, last_row: 17 },
      { header_row: 24, last_row: 58 }
    ]
  },{
    name: 'NCCIS_10-11_to_16-17', 
    data_blocks: [
      { header_row: 5, last_row: 27 },
      { header_row: 35, last_row: 76 }
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

sheets_to_process.each do |sheet|
  puts sheet[:name]
  sheet[:data_blocks].each do |block|
    block_name = block[:table_name] || sheet[:name]
    puts "\t#{block_name}: #{data_tables_workbook.sheet(sheet[:name]).row(block[:header_row])[0]}"
  end
  
  #puts sheet.row(5).cell(1)
  #puts "---------"
end
