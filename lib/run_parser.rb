require_relative 'animal_parser'

dogs_url = 'https://schroniskochorzow.pl/psy/'
cats_url = 'https://schroniskochorzow.pl/koty/'
new_arrivals_url = 'https://schroniskochorzow.pl/nowo-przyjete/'

dogs_parser = AnimalParser.new(dogs_url)
cats_parser = AnimalParser.new(cats_url)
new_arrivals_parser = AnimalParser.new(new_arrivals_url)

dogs_data = dogs_parser.parse_dogs
cats_data = cats_parser.parse_cats
new_arrivals_data = new_arrivals_parser.parse_new_arrivals

dogs_parser.generate_excel(dogs_data, cats_data, new_arrivals_data)
