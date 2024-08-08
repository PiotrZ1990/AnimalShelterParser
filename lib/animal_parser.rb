require 'nokogiri'
require 'httparty'
require 'write_xlsx'

class AnimalParser
  def initialize(base_url)
    @base_url = base_url
  end

  def fetch_data(url)
    response = HTTParty.get(url)
    Nokogiri::HTML(response.body)
  end

  def parse_dogs
    page_number = 1
    animals = []

    loop do
      url = "#{@base_url}/psy/page/#{page_number}/"
      doc = fetch_data(url)
      items = doc.css('.rt-grid-item')

      break if items.empty?

      items.each do |animal|
        name = animal.css('h2 strong').text.strip
        details = animal.css('p').text.strip
        image_url = animal.css('.rt-img-holder img').attr('src')&.value

        # Extracting information using regex
        number = details.match(/Numer:\s*([\d\/-]+)/)&.captures&.first&.strip
        gender = details.match(/Płeć:\s*(samiec|samica|samce|samice)/)&.captures&.first&.strip
        birth_year_match = details.match(/Wiek:\s*ur\.\s*(?:\d{2}\.)?\s*(\d{4})/)
        birth_year = birth_year_match ? "ur. #{birth_year_match[1]}" : nil
        size = details.match(/Rozmiar:\s*(.*)/)&.captures&.first&.strip

        animals << {
          name: name,
          number: number,
          gender: gender,
          age: birth_year,
          size: size,
          image_url: image_url,
          species: 'Pies'
        }
      end

      page_number += 1
    end

    animals
  end

  def parse_cats
    page_number = 1
    animals = []

    loop do
      url = "#{@base_url}/koty/page/#{page_number}/"
      doc = fetch_data(url)
      items = doc.css('.rt-grid-item')

      break if items.empty?

      items.each do |animal|
        name = animal.css('h2 strong').text.strip
        details = animal.css('p').text.strip
        image_url = animal.css('.rt-img-holder img').attr('src')&.value

        # Extracting information using regex
        number = details.match(/Numer:\s*([\d\/-]+)/)&.captures&.first&.strip
        gender = details.match(/Płeć:\s*(samiec|samica)/)&.captures&.first&.strip
        age = details.match(/Wiek:\s*(.*?)\s*(Znaleziona|Znaleziony|$)/)&.captures&.first&.strip
        fiv_felv_test_match = details.match(/Test FIV\/FELV\s*\([^\)]+\)|Test FIV\s*\([^\)]+\)\s*\/\s*FELV\s*\([^\)]+\)/)
        fiv_felv_test = fiv_felv_test_match ? fiv_felv_test_match[0].strip : 'Brak testu'

        animals << {
          name: name,
          number: number,
          gender: gender,
          age: age,
          fiv_felv_test: fiv_felv_test,
          image_url: image_url,
          species: 'Kot'
        }
      end

      page_number += 1
    end

    animals
  end

  def parse_new_arrivals
    url = "#{@base_url}/nowo-przyjete/"
    doc = fetch_data(url)
    animals = []

    doc.css('.rt-grid-item').each do |animal|
      number = animal.css('h2 strong').text.strip
      details = animal.css('p').text.strip
      image_url = animal.css('.rt-img-holder img').attr('src')&.value

      # Extracting information using regex
      quarantine_until = details.match(/Kwarantanna do:\s*(\d{2}\.\d{2}\.\d{4})/)&.captures&.first&.strip || 'Brak daty'
      gender = details.match(/Płeć:\s*(samiec|samica)/)&.captures&.first&.strip || 'Brak płci'
      age = details.match(/Wiek:\s*(.*?)\s*(Znaleziona|Znaleziony|$)/)&.captures&.first&.strip || 'Brak wieku'
      found = details.match(/(Znaleziona|Znaleziony):\s*(.*)/)&.captures&.last&.strip || 'Brak miejsca'

      animals << {
        name: number,
        number: number,
        gender: gender,
        age: age,
        found: found,
        image_url: image_url,
        species: 'Nowo przybyłe'
      }
    end

    animals
  end

  def generate_excel(all_data)
    workbook = WriteXLSX.new('Schronisko Chorzow.xlsx')
    sheet = workbook.add_worksheet('Zwierzęta')

    sheet.write_row(0, 0, ["Imię", "Numer", "Płeć", "Wiek", "Rozmiar", "Testy", "Znaleziona", "Gatunek", "URL zdjęcia"])

    all_data.each_with_index do |animal, index|
      sheet.write_row(index + 1, 0, [
        animal[:name], animal[:number], animal[:gender], animal[:age], animal[:size], 
        animal[:fiv_felv_test], animal[:found], animal[:species], animal[:image_url]
      ])
    end

    workbook.close
  end

  def parse_all
    dogs = parse_dogs
    cats = parse_cats
    new_arrivals = parse_new_arrivals

    generate_excel(dogs + cats + new_arrivals)
  end
end

# Użycie parsera
base_url = 'https://schroniskochorzow.pl'
parser = AnimalParser.new(base_url)
parser.parse_all