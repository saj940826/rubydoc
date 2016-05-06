require 'docx'

doc =  Docx::Document.open('Buried Infrastructure Performance 2015-07-15.docx')
title = doc.paragraphs[0]
doc.paragraphs.delete(0)
#projects split
doc.paragraphs.each_with_index do |paragraph, count|
  if paragraph.to_s =~ /^\s*[A-Z]\..*\(.*\-.*\)\s*$/

  end
end

categoryN = Array.new
categoryName = Array.new
#category split
doc.paragraphs.each_with_index do |paragraph, count|
  if paragraph.to_s =~ /^\s*(I|II|III)\..*(Know|Eco|Tech).*/
    categoryName.push paragraph
    categoryN.push count
  end
end
puts doc.paragraphs[categoryN[1]]
categoryProjects = Array.new
categoryProjects.push doc.paragraphs[categoryN[0]..categoryN[1]-1]
puts categoryProjects.compact