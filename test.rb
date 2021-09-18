
require 'puredocx'
require 'docxgen'

PureDocx.create('./example1.docx', paginate_pages: 'right') do |doc|
  doc.content([doc.text('text', style: [:bold], size: 32, align: 'center')])
end

generator = Docxgen::Generator.new("template.docx")
generator.render({ variable: "Value" }, remove_missing: true)
generator.valid? # Returns true if there is no errors
puts generator.errors # Returns array of the errors
generator.save("result.docx")
