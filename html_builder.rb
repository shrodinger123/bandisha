require 'rubyXL'
require 'pry'
puts "Loading products ..."
# workbook = RubyXL::Parser.parse 'ynf_products/Merge (14-08-2024).xlsx'
workbook = RubyXL::Parser.parse 'ynf_products/sarees-5001-end.xlsx'
# workbook = RubyXL::Parser.parse 'ynf_products/test.xlsx'
# workbook = RubyXL::Parser.parse 'ynf_products/OnlySarees.xlsx'
puts "Workbook loaded .."
worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"

worksheet1 = worksheets[0]

@start_row = 2
@end_row = 380

File.open("ynf_products/sarees-5001-end.html", 'w') do |f|
  f.write(%Q[<!DOCTYPE html>
<html>
<head>
    <style>
        table, th, td {
            border: 1px solid black;
        }
        img {
            display: block;
            max-width:200px;
            max-height:300px;
            width: auto;
            height: auto;
        }
    </style>
</head>
<body>
<h1>Sarees</h1>
<table>
    <tr>
        <th>Product Code</th>
        <th>Product Variant Code</th>
        <th>Cover Image</th>
        <th>Image1</th>
        <th>Image2</th>
        <th>Image3</th>
        <th>Image4</th>
        <th>Image5</th>
    </tr>])
  (@start_row..@end_row).each do |row|
    if  worksheet1.cell_at("H#{row}")&.value&.include?("Saree") &&  worksheet1.cell_at("Q#{row}")&.value&.to_i >= 20 &&  worksheet1.cell_at("AO#{row}")&.value&.include?("Saree")
      product_code_cell = worksheet1.cell_at("B#{row}")&.value
      product_variant_code_cell = worksheet1.cell_at("C#{row}")&.value
      cover_image_cell = worksheet1.cell_at("Z#{row}")&.value
      image1_cell = worksheet1.cell_at("AB#{row}")&.value
      image2_cell = worksheet1.cell_at("AD#{row}")&.value
      image3_cell = worksheet1.cell_at("AF#{row}")&.value
      image4_cell = worksheet1.cell_at("AH#{row}")&.value
      image5_cell = worksheet1.cell_at("AJ#{row}")&.value
      # puts "#{product_code_cell}"
      # puts "#{product_variant_code_cell}"
      # puts "#{cover_image_cell}"
      # puts "#{image1_cell}"
      # puts "#{image2_cell}"
      # puts "#{image3_cell}"
      # puts "#{image4_cell}"
      # puts "#{image5_cell}"
      f.write( %Q[
        <tr>
        <td>#{product_code_cell}</td>
        <td>#{product_variant_code_cell}</td>
        <td><img src="#{cover_image_cell}"></td>
        <td><img src="#{image1_cell}"></td>
        <td><img src="#{image2_cell}"></td>
        <td><img src="#{image3_cell}""></td>
        <td><img src="#{image4_cell}""></td>
        <td><img src="#{image5_cell}""></td>
        </tr>])
      end
  end
  f.write(%Q[
      </table>
      </body>
      </html>
          ])
end

