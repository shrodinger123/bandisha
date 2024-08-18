require 'rubyXL'
require 'csv'
require 'pry'
require 'pp'

TEMPLATE_SHOPIFY_CSV = "shopify_template.csv"
YNF_PRODUCT_XLSX = "ynf_template.xlsx"
GENERATED_SHOPIFY_CSV ="generated_shopify2.csv"


# TAGS = [material,price_range, style, celebrity, "trending", "wedding-gift"]
# shopify_csv = CSV.read(TEMPLATE_SHOPIFY_CSV)
shopify_csv = CSV.parse(File.read(TEMPLATE_SHOPIFY_CSV), headers: true)
# puts "Loading products ..."
@ynf_wb = RubyXL::Parser.parse YNF_PRODUCT_XLSX
@ynf_ws = @ynf_wb.worksheets[0]

def generate_product_code(row)
  @ynf_ws.cell_at("B#{row}")&.value
end

def generate_product_variant_code(row)
  @ynf_ws.cell_at("C#{row}")&.value
end

def generate_title(row)
  @ynf_ws.cell_at("A#{row}")&.value
end

def generate_body_html(row)
  "<p> #{@ynf_ws.cell_at("J#{row}")&.value} </p>"
end

def generate_tags(row, pkg)
  all_tags = []

  # price tags = [lt_1000, bt_1000_2000, bt_2000_3000, bt_3000_4000, bt_4000_6000, bt_6000_10000, gt_10000]
  price = generate_price(row, pkg)
  price_tag = ''

  if price < 1000
    price_tag = "lt_1000"
  elsif price >= 1000 && price <=2000
    price_tag = "bt_1000_2000"
  elsif price > 2000 && price <=3000
    price_tag = "bt_2000_3000"
  elsif price > 3000 && price <=4000
    price_tag = "bt_3000_4000"
  elsif price > 4000 && price <=6000
    price_tag = "bt_4000_6000"
  elsif price > 6000 && price <10000
    price_tag = "bt_6000_10000"
  elsif price >= 10000
    price_tag = "gt_10000"
  end

  all_tags << price_tag
  all_tags << "trending"
  all_tags << "wedding-gift"
  t_or_w =  @ynf_ws.cell_at("BG#{row}")&.value
  all_tags.join(",")


end

def generate_price(row, pkg)
  b2b_discounted_price = @ynf_ws.cell_at("P#{row}")&.value&.to_i
  pkg_price = (pkg=="basic" ? 0 : 20)
  selling_price = b2b_discounted_price * 1.5 * 1.05 + pkg_price
  selling_price
end

def generate_color_value(row)
  @ynf_ws.cell_at("AR#{row}")&.value
end

def generate_packaging_value(pkg)
  if pkg=="basic"
    return "Basic (+ Rs. 0)"
  elsif pkg =="premium"
    return "Premium (+ Rs.20)"
  end
end

def generate_variant_sku(row)
  @ynf_ws.cell_at("C#{row}")&.value
end

def generate_variant_grams(row)
  @ynf_ws.cell_at("Y#{row}")&.value
end

def generate_variant_inventory_quantity(row)
  @ynf_ws.cell_at("Q#{row}")&.value
end

def generate_variant_compare_at_price(row,pkg)
  generate_price(row, pkg)
end

def generate_image_src(row)
  @ynf_ws.cell_at("Z#{row}")&.value
end

product_hash = {}

CSV.open(GENERATED_SHOPIFY_CSV, "w") do |csv|
  header_row = CSV.read(TEMPLATE_SHOPIFY_CSV)[0]
  csv << header_row
  (2..100).each do |row|
    if  @ynf_ws.cell_at("H#{row}")&.value&.include?("Saree") &&  @ynf_ws.cell_at("Q#{row}")&.value&.to_i >= 20 &&  @ynf_ws.cell_at("AO#{row}")&.value&.include?("Saree") && !@ynf_ws.cell_at("BG#{row}").nil?
      puts "Selected : #{@ynf_ws.cell_at("C#{row}").value}"
      product_code = generate_product_code(row)
      product_variant_code = generate_product_variant_code(row)
      image_position = 0
      #generate variant hash per packaging
      ["basic", "premium"].each do |pkg|
        new_sku = false
        if !product_hash.key?(product_code)
          product_hash[product_code] = []
          new_sku = true
          image_position = 0
        end
        variant_hash = {
          "Handle" => generate_product_code(row),
          "Title" => new_sku ? generate_title(row) : "",
          "Body (HTML)" => new_sku ? generate_body_html(row) : "",
          "Vendor" => new_sku ? "Bandisha" : "",
          "Product Category" => new_sku ? "Apparel & Accessories > Clothing > Traditional & Ceremonial Clothing > Saris & Lehengas" : "",
          "Type" => "",
          "Tags" => generate_tags(row,pkg),
          "Published" => "true",
          "Option 1 Name" => new_sku ? "Color" : "",
          "Option 1 Value" => generate_color_value(row),
          "Option 1 Linked To" => "",
          "Option 2 Name" => new_sku ? "Packaging Options" : "",
          "Option 2 Value" => generate_packaging_value(pkg),
          "Option 2 Linked To" => "",
          "Option 3 Name" => "",
          "Option 3 Value" => "",
          "Option 3 Linked To" => "",
          "Variant SKU" => generate_variant_sku(row),
          "Variant Grams" => generate_variant_grams(row),
          "Variant Inventory Tracker" => "Shopify",
          "Variant Inventory Quantity" => generate_variant_inventory_quantity(row),
          "Variant Inventory Policy" => "deny",
          "Variant Fulfillment Service" => "manual",
          "Variant Price" => generate_price(row, pkg),
          "Variant Compare At Price" => generate_variant_compare_at_price(row, pkg),
          "Variant Requires Shipping" => "true",
          "Variant Taxable" => "false",
          "Variant Barcode" => nil,
          "Image Src" => generate_image_src(row),
          "Image Position" => nil,
          "Image Alt Text" => "",
          "Gift Card" => "false",
          "SEO Title" => nil,
          "SEO Description" => nil,
          "Google Shopping / Google Product Category" => nil,
          "Google Shopping / Gender" => nil,
          "Google Shopping / Age Group" => nil,
          "Google Shopping / MPN" => nil,
          "Google Shopping / Condition" => nil,
          "Google Shopping / Custom Product" => nil,
          "Google Shopping / Custom Label 0" => nil,
          "Google Shopping / Custom Label 1" => nil,
          "Google Shopping / Custom Label 2" => nil,
          "Google Shopping / Custom Label 3" => nil,
          "Google Shopping / Custom Label 4" => nil,
          "Complementary products (product.metafields.shopify--discovery--product_recommendation.complementary_products)" => nil,
          "Related products (product.metafields.shopify--discovery--product_recommendation.related_products)" => nil,
          "Related products settings (product.metafields.shopify--discovery--product_recommendation.related_products_display)" => nil,
          "Variant Image" => nil,
          "Variant Weight Unit" => "g",
          "Variant Tax Code" => nil,
          "Cost per item" => nil,
          "Included / India" => "true",
          "Price / India" => nil,
          "Compare At Price / India" => nil,
          "Included / International" => "true",
          "Price / International" => nil,
          "Compare At Price / International" => nil,
          "Status" => "active"
        }
        csv_row = variant_hash.values
        csv << csv_row
      end
      product_hash[product_code] << product_variant_code
    end
  end
end





# product_code_cell = worksheet1.cell_at("B#{row}")&.value

# to do
# Add packaging options and generate price
# variant price = B2B price * 1.5 * 1.05 + packaging cost (20 for premium)