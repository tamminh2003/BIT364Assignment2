#            _____ _____ _____ _____ _   _ __  __ ______ _   _ _______   ___    _____        _____ _______   ___
#     /\    / ____/ ____|_   _/ ____| \ | |  \/  |  ____| \ | |__   __| |__ \  |  __ \ /\   |  __ \__   __| |__ \
#    /  \  | (___| (___   | || |  __|  \| | \  / | |__  |  \| |  | |       ) | | |__) /  \  | |__) | | |       ) |
#   / /\ \  \___ \\___ \  | || | |_ | . ` | |\/| |  __| | . ` |  | |      / /  |  ___/ /\ \ |  _  /  | |      / /
#  / ____ \ ____) |___) |_| |_ |__| | |\  | |  | | |____| |\  |  | |     / /_  | |  / ____ \| | \ \  | |     / /_
# /_/    \_\_____/_____/|_____\_____|_| \_|_|  |_|______|_| \_|  |_|    |____| |_| /_/    \_\_|  \_\ |_|    |____|

# by Tam Minh TRAN - S1499037

# List of dependencies
require 'tty-prompt' # Get user input
require 'figlet' # Fancy title
require 'search_in_file' # Search for term in .docx file
require 'file-find' # Search for .docx file
require 'puredocx' # Generate docx

class Assignment2
  # Initilize/construct method for class Assignment 2
  def initialize
    @font = Figlet::Font.new('big.flf') # Setup class variable @font
    @figlet = Figlet::Typesetter.new(@font) # Setup figlet - fancy title generator
    @prompt = TTY::Prompt.new # Setup fancy prompt using tty-prompt
  end

  # This method uses figlet gem to generate a fancy title
  def display_title
    puts @figlet['ASSIGNMENT 2 PART 2']
    puts "by Tam Minh TRAN - S1499037\n\n"
  end

  # This method uses tty-prompt gem to ask user for input
  # The advantages of using tty-prompt gem is the input from user is processed
  # with the removal of trailing whitespaces and control symbols.
  # Hence, tty-prompt gem assist developers with a reliable prompt interface.
  def read_input
    puts '========================'
    # tty-prompt's ask method gets search term from user and assign to @term
    @term = @prompt.ask('What are you searching for?')

    puts '========================'
    # tty-prompt's ask method gets filepath from user and assign to @path
    @path = @prompt.ask('Where are you searching in?')

    puts '========================'
    # Display received input for @path, in case user typed incorrectly.
    puts "Searching for .docx files in #{@path}"
  end

  # This method uses find-file gem to search for .docx files in given filepath.
  # Using find-file gem saved time comparing using Ruby built-in API.
  # find-file's find method returns an array of filepath matching the defined rules.
  def search_file
    # This defines the rules used for searching the .docx files
    @rule = File::Find.new(
      pattern: '*.docx',
      follow: false,
      path: [@path],
      maxdepth: 1 # maxdepth is used to define find-file only search for current folder, and not the sub folders
    )

    # This runs the find method of find-file gem and assign returned array of filepath to @docx_file
    @docx_file = @rule.find

    # In case of no .docx files found, puredocx's create method is invoked to create a .docx file containing the term.
    if @docx_file.empty?
      puts '========================'
      puts "No .docx files in folder => Creating new #{@term}.docx file"
      PureDocx.create("./#{@term}.docx", paginate_pages: 'right') do |doc|
        # puredocx's create method is executed here
        doc.content([doc.text(@term, style: [:bold], size: 32, align: 'center')])
      end
      exit # Stop the program once document is created.
    end
  end

  # This method uses search_in_file gem to search for the input @term
  # The search method returns an array of hashes (structured: file, paragraphs)
  def search_term
    @found_file = [] # Clear @found_file, this is where we keep the list of files containing @term.
    @docx_file.each do |f|
      # Iterate through .docx files found in the @path
      rel_hash = SearchInFile.search(f, @term) # Search the file for @term, then store in rel_hash
      @found_file << rel_hash[0] if rel_hash[0] # If rel_hash[0] is not empty, push search result to @found_file
    end

    # This executes puredocx to creat .docx file if there are no files in the folder containing @term
    if @found_file.empty? || @found_file.size.zero?
      puts '========================'
      puts "Cannot find files with term '#{@term}' => Creating new #{@term}.docx file"
      PureDocx.create("./#{@term}.docx", paginate_pages: 'right') do |doc|
        # puredocx's create method is executed here
        doc.content([doc.text(@term, style: [:bold], size: 32, align: 'center')])
      end
      exit
    end
  end

  # This method displays the final result of the search
  def display_result
    puts '========================'
    puts "Found #{@found_file.size} files with term '#{@term}'"
    @found_file.each do |f| # Iterate through array of files containing @term.
      puts "=> #{f[:file]}; found #{f[:paragraphs][0].scan(@term).size} occurrence(s)"
      # :file containing the path to the file
      # NOTE: On Windows system, the end of line is \r\n.
      # However, search_in_file detects paragraph using UNIX end of line symbols \n.
      # Hence, the :paragraph returned is whole document instead of seperated paragraphs as intended.
      # To count the occurrences of @term within the document, String's scan method is used.
      # scan method returns an array of matches with @term.
      # Then using size method of Array to determine the number of occurrences
    end
  end
end

# ==== PROGRAM MAIN PROCESS ==== #
assignment2 = Assignment2.new # Create new instance of Assignment2
assignment2.display_title # Display fancy title
assignment2.read_input # Read user input
assignment2.search_file # Search for docx files
assignment2.search_term # Search for term in files
assignment2.display_result # Display results

# TODO: - ERROR HANDLING -- File with the same name with what @term.docx, and other errors.
