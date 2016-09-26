require 'roo'
require 'csv'
require 'date'

begin
  print 'Enter .xlsx file name: '
  file_name = gets

  xlsx = Roo::Excelx.new(file_name.chomp)
  csv = CSV.open('output.csv', 'w')

  xlsx.each_with_pagename do |name, sheet|
    puts "Parsing sheet #{name}"

    current_date = nil

    (1..sheet.last_row).each do |row|
      a_col = sheet.cell(row, 'A')

      next unless a_col.is_a?(Date) || (a_col.is_a?(String) && !a_col.empty?)

      if a_col.is_a?(Date)
        current_date = a_col
        puts "Parsing: #{a_col}"
        next
      end

      next unless a_col.split(':').size == 2 && a_col =~ /([0-9]{2}:[0-9]{2})/

      csv_row = [current_date.to_s]
      ('A'..'AQ').each do |col|
        current_cell = sheet.cell(row, col)
        next unless current_cell
        csv_row << current_cell
      end

      csv << csv_row
    end
  end

  puts 'output.csv generated.'

  csv.close
  xlsx.close

rescue => e
  puts "Error: #{e}"
end
