require 'nokogiri'
require 'httparty'
require 'write_xlsx'

class AnimalParser
  def initialize(url)
    @url = url
  end

  def fetch_data
    response = HTTParty.get(@url)
    Nokogiri::HTML(response.body)
  end

def parse_dogs
  doc = fetch_data
  animals = []

  doc.css('.rt-grid-item').each do |animal|
    name = animal.css('h2 strong').text.strip
    details = animal.css('p').text.strip

    # Debugging output
    puts "Dog details: #{details.inspect}"

    # Extracting information using regex
    number = details.match(/Numer:\s*([\d\/-]+)/)&.captures&.first&.strip
    gender = details.match(/Płeć:\s*(samiec|samica|samce|samice)/)&.captures&.first&.strip

    # Extract birth year and month (if available)
    birth_year_match = details.match(/Wiek:\s*ur\.\s*(?:\d{2}\.)?\s*(\d{4})/)
    if birth_year_match
      year = birth_year_match[1]   # Extract year
      birth_year = "ur. #{year}"   # Format as "ur. YYYY"
    else
      birth_year = nil
    end

    # Size (Rozmiar) extraction, if applicable
    size = details.match(/Rozmiar:\s*(.*)/)&.captures&.first&.strip

    animals << {
      name: name,
      number: number,
      gender: gender,
      birth_year: birth_year,
      size: size
    }
  end

  animals
end






  def parse_cats
    doc = fetch_data
    animals = []

    doc.css('.rt-grid-item').each do |animal|
      name = animal.css('h2 strong').text.strip
      details = animal.css('p').text.strip

      # Debugging output
      puts "Cat details: #{details.inspect}"

      # Extracting information using regex
      number = details.match(/Numer:\s*([\d\/-]+)/)&.captures&.first&.strip
      gender = details.match(/Płeć:\s*(samiec|samica)/)&.captures&.first&.strip
      age = details.match(/Wiek:\s*(ok\.\s*\d+\s*(tyg\.*|mies\.*|lat))/)&.captures&.first&.strip

      # Extract FIV/FELV test result with various possible formats
      fiv_felv_test_match = details.match(/Test FIV\/FELV\s*\([^\)]+\)|Test FIV\s*\([^\)]+\)\s*\/\s*FELV\s*\([^\)]+\)/)
      fiv_felv_test = fiv_felv_test_match ? fiv_felv_test_match[0].strip : 'Brak testu'

      animals << {
        name: name,
        number: number,
        gender: gender,
        age: age,
        fiv_felv_test: fiv_felv_test
      }
    end

    animals
  end



def parse_new_arrivals
    doc = fetch_data
    animals = []

    doc.css('.rt-grid-item').each do |animal|
      # Extract number from <h2> element
      number = animal.css('h2 strong').text.strip

      # Extract details from <p> element
      details = animal.css('p').text.strip

      # Debugging output
      puts "New arrival details: #{details.inspect}"
      puts "Number: #{number.inspect}"

      # Wyciąganie informacji przy użyciu regex
      quarantine_until = details.match(/Kwarantanna do:\s*(\d{2}\.\d{2}\.\d{4})/)&.captures&.first&.strip || 'Brak daty'
      gender = details.match(/Płeć:\s*(samiec|samica)/)&.captures&.first&.strip || 'Brak płci'
      age = details.match(/Wiek:\s*(.*?)\s*(Znaleziona|Znaleziony|$)/)&.captures&.first&.strip || 'Brak wieku'
      found = details.match(/(Znaleziona|Znaleziony):\s*(.*)/)&.captures&.last&.strip || 'Brak miejsca'

      # Debugging output for extracted fields
      puts "Quarantine Until: #{quarantine_until.inspect}"
      puts "Gender: #{gender.inspect}"
      puts "Age: #{age.inspect}"
      puts "Found: #{found.inspect}"

      # Dodawanie do listy zwierząt
      animals << {
        number: number,
        quarantine_until: quarantine_until,
        gender: gender,
        age: age,
        found: found
      }
    end

    animals
  end







  def generate_excel(dogs_data, cats_data, new_arrivals_data)
    workbook = WriteXLSX.new('animals.xlsx')

    # Dogs sheet
    dogs_sheet = workbook.add_worksheet('Psy')
    dogs_sheet.write_row(0, 0, ["Imię", "Numer", "Płeć", "Wiek", "Rozmiar"])
    dogs_data.each_with_index do |animal, index|
      dogs_sheet.write_row(index + 1, 0, [animal[:name], animal[:number], animal[:gender], animal[:birth_year], animal[:size]])
    end

    # Cats sheet
    cats_sheet = workbook.add_worksheet('Koty')
    cats_sheet.write_row(0, 0, ["Imię", "Numer", "Płeć", "Wiek", "Testy"])
    cats_data.each_with_index do |animal, index|
      cats_sheet.write_row(index + 1, 0, [animal[:name], animal[:number], animal[:gender], animal[:age], animal[:fiv_felv_test]])
    end

    # New arrivals sheet
    new_arrivals_sheet = workbook.add_worksheet('Nowo przyjęte')
    new_arrivals_sheet.write_row(0, 0, ["Numer", "Kwarantanna do:", "Płeć", "Wiek", "Znaleziona"])
    new_arrivals_data.each_with_index do |animal, index|
      new_arrivals_sheet.write_row(index + 1, 0, [animal[:number], animal[:quarantine_until], animal[:gender], animal[:age], animal[:found]])
    end

    workbook.close
  end
end
