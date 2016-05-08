require 'docx'
# document convert
def docx_to_hashs path
  @path = File.absolute_path path
  doc = Docx::Document.open(@path)
  document = doc.paragraphs
  portfolio_title = get_portfolio_title(document)
  all_title_with_line, all_code = get_project_titles_with_code document
  all_category_with_line = get_category document
  all_key = Array.new
  all_value = Array.new
  all_projects_title = Array.new
  all_projects_category = Array.new
  #get key and value

  doc.tables.each do |table|
    key = Array.new
    mutiple_value = Array.new
    table.rows.each do |row|
      value = Array.new
      key.push(row.cells[0].text)
      for column in 1..row.cells.count-1 do
        value.push row.cells[column].text
      end
      mutiple_value.push value
    end
    all_key.push key
    all_value.push mutiple_value
  end
  all_projects_hashes = Array.new
  all_key.zip(all_value).each do |key, value|
    hash = {}
    key.zip(value) { |a, b| hash[a.to_sym] = b }
    number = nil
    all_code.each_with_index do |codes, count|
      codes.each do |code|
        if hash["Project number".to_sym].first.to_s.include? code
          number = count
        end
      end
    end
    all_projects_category.push (number.nil? ? nil : (get_project_category all_title_with_line[number], all_category_with_line))
    all_projects_title.push (number.nil? ? nil : all_title_with_line[number].keys.first)
    all_projects_hashes.push hash
  end
  #type   title category  extra
  return portfolio_title, all_projects_title, all_projects_category, all_projects_hashes
end

def get_portfolio_title document
  document.each do |paragraph|
    if paragraph.to_s =~ /^\s*$/ || paragraph.to_s == ''
    else
      return paragraph.to_s
    end
  end
end

#Method for convert docx to hashes
def hashs_to_one_text hashs
  text = String.new
  hashs.each do |key, values|
    text << '|$$|'
    text << key.to_s
    text << '|@@|'
    values.each do |value|
      text << value.to_s
      if value != values.last
        text << '|##|'
      end
    end
  end
  return text
end

def text_to_hashes text
  text_array = text.split('|$$|').reject { |t| t.empty? }
  all_key = Array.new
  all_value = Array.new
  text_array.each do |rowHash|
    key_textValue = rowHash.split('|@@|')
    all_key.push key_textValue[0]
    values = key_textValue[1].split('|##|')
    all_value.push values
  end
  extra_hashes = Hash[all_key.zip(all_value)]
  return extra_hashes
end

def get_project_titles_with_code paragraphs
  all_title_with_line = Array.new
  all_code = Array.new
  paragraphs.each_with_index do |paragraph, count|
    if paragraph.to_s =~ /^\s*[A-Z]\..*\(.*\-.*\)\s*$/
      all_title_with_line.push Hash[paragraph.to_s.gsub(/\s*\(.*\-.*\)\s*$/, '').gsub(/^\s*[A-Z]\.\s*/, ''), count]
      row_code = /\(.*\-.*\)/.match(paragraph.to_s).to_s
      all_code.push /(?!\().*\-.*(?=\))/.match(row_code).to_s.delete(' ').split(',')
    end
  end
  return all_title_with_line, all_code
end

def get_category paragraphs
  all_category_with_line = Array.new
  #category split
  paragraphs.each_with_index do |paragraph, count|
    if paragraph.to_s =~ /^\s*(I|II|III|IV|V)\..*(Know|Eco|Tech).*/
      all_category_with_line.push Hash[paragraph.to_s.sub(/^\s*(I|II|III|IV|V)\.\s*/, ''), count]
    end
  end
  return all_category_with_line
end

def get_project_category title_with_line, all_category_with_line
  category = nil
  all_category_with_line.each do |category_wih_line|
    if category_wih_line.values.first < title_with_line.values.first
      category = 1 if category_wih_line.keys.first.to_s.upcase.include?("ECONO")
      category = 2 if category_wih_line.keys.first.to_s.upcase.include?("KNOWLE")
      category = 3 if category_wih_line.keys.first.to_s.upcase.include?("TECHNOLO") else 4
    end
  end
  return category
end

puts docx_to_hashs 'Buried Infrastructure Performance 2015-07-15.docx'