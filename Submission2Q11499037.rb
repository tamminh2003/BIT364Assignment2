require 'tty-prompt'
require 'figlet'
require 'search_in_file'
require 'file-find'
require 'puredocx'

font = Figlet::Font.new('big.flf')

figlet = Figlet::Typesetter.new(font)

puts figlet['ASSIGNMENT 2 PART 2']
puts "by Tam Minh TRAN - S1499037\n\n"

prompt = TTY::Prompt.new

puts '========================'
term = prompt.ask('What are you searching for?')

puts '========================'
path = prompt.ask('Where are you searching in?')

puts '========================'
puts "Searching for .docx files in #{path}"

rule = File::Find.new(
  pattern: '*.docx',
  follow: false,
  path: [path],
  maxdepth: 1
)

found_file = []
docx_file = rule.find

if docx_file.empty?
  puts '========================'
  puts "No .docx files in folder => Creating new #{term}.docx file"
  PureDocx.create("./#{term}.docx", paginate_pages: 'right') do |doc|
    doc.content([doc.text(term, style: [:bold], size: 32, align: 'center')])
  end
  exit
end

docx_file.each do |f|
  rel_hash = SearchInFile.search(f, term)
  found_file << rel_hash[0] if rel_hash[0]
end

if found_file.empty? || found_file.size.zero?
  puts '========================'
  puts "Cannot find files with term '#{term}' => Creating new #{term}.docx file"
  PureDocx.create("./#{term}.docx", paginate_pages: 'right') do |doc|
    doc.content([doc.text(term, style: [:bold], size: 32, align: 'center')])
  end
  exit
end

puts '========================'
puts "Found #{found_file.size} files with term '#{term}'"
found_file.each do |f|
  puts "=> #{f[:file]}; found #{f[:paragraphs][0].scan(term).size} occurrence(s)"
end

