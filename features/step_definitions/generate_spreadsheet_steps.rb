require 'roo'

Given /^I generate following spreadsheet:$/ do |table|
  @excel = LazyWombat::Xlsx.new do
    row do
      table.hashes.first.keys.each { |key| cell key }
    end
    table.hashes.each do |hash|
      row do
        hash.values.each { |value| cell value }
      end
    end
  end
end

Given /^save it as "(.*?)"$/ do |file_name|
  @excel.save file_name
end

Then /^"(.*?)" file should exist$/ do |file_name|
  File.exists? file_name
end

Then /^"(.*?)" should have (\d+) spreadsheet$/ do |file_name, count|
  Excelx.new(file_name).sheets.should have(count).records
end

# gem 'roo' casts all read numbers to float
Then /^"(.*?)" should look like this:$/ do |file_name, table|
  workbook = Excelx.new file_name
  prototype = table.raw
  columns_count = prototype.first.count
  1.upto(workbook.last_row) do |row_number|
    row = columns_count.times.map { |column| workbook.cell row_number, column + 1 }
    row.should == row_with_types(prototype[ row_number - 1 ])
  end
end

def row_with_types row
  return if row.nil?
  row.each_with_index do |value, index|
    row[index] = (value =~ /\<(.*)\>/) ? eval(value[1..-2]) : value
  end
end