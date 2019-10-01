import urllib.request
import openpyxl
import csv
import re
import os

URL = 'https://www.uke.gov.pl/pozwolenia-radiowe-dla-klasycznych-sieci-radiokomunikacji-ruchomej-ladowej-5458'
REGEX = '<a href="/files/\?id_file=([0-9]+)">[a-zA-Z0-9_-]+\.xlsx<\/a>'
PREFIX = 'https://www.uke.gov.pl/files/?id_file='

def download_file_list():
	connection = urllib.request.urlopen(URL)
	charset = connection.info().get_content_charset()
	wegpage_contents = connection.read().decode(charset)
	connection.close()

	file_list = []
	for file_id in re.findall(REGEX, wegpage_contents):
		file_list.append(PREFIX + file_id)

	return file_list

def save_headers(csv_writer):
	headers = []
	headers.append('Nr pozwolenia')
	headers.append('Data wygaśnięcia')
	headers.append('Nazwa stacji')
	headers.append('Rodzaj stacji')
	headers.append('Rodzaj sieci')
	headers.append('Długość geograficzna')
	headers.append('Szerokość geograficzna')
	headers.append('Promień obszaru obsługi')
	headers.append('Lokalizacja stacji')
	headers.append('ERP')
	headers.append('Azymut')
	headers.append('Elewacja')
	headers.append('Polaryzacja')
	headers.append('Zysk anteny')
	headers.append('Wysokość umieszczenia anteny')
	headers.append('Wysokość terenu')
	headers.append('Charakterystyka promieniowania - poziom')
	headers.append('Charakterystyka promieniowania - pion')
	headers.append('Częstotliwości nadawcze')
	headers.append('Częstotliwości odbiorcze')
	headers.append('Szerokości kanałów nadawczych')
	headers.append('Szerokości kanałów odbiorczych')
	headers.append('Operator')
	headers.append('Adres operatora')
	csv_writer.writerow(headers)

def process_file(csv_writer, file_name):
	folder = openpyxl.load_workbook(file_name)
	for sheet in folder:
		first_row = True
		for row in sheet:
			if not first_row:
				process_row(csv_writer, row)
			else:
				first_row = False

def process_row(csv_writer, row):
	row_for_csv = []
	for cell in row:
		row_for_csv.append(cell.value)
	row_for_csv[5] = fix_coords(row_for_csv[5])
	row_for_csv[6] = fix_coords(row_for_csv[6])
	csv_writer.writerow(row_for_csv)

def fix_coords(coord):
	if not coord: return ''
	coord = re.split('E|N|\'|"', coord)
	return int(coord[0]) + int(coord[1])/60 + int(coord[2])/3600

def main():
	print('Otwieranie pliku db.csv...')
	with csv_file = open('db.csv', 'w', newline='', encoding='utf-8')
		csv_writer = csv.writer(csv_file)

		print('Dopisywanie wiersza nagłówkowego do pliku...')
		save_headers(csv_writer)

		for file_ in download_file_list():
			print('Przetwarzanie pliku ' + file_ + '...')
			urllib.request.urlretrieve(file_, 'tmp.xlsx')
			process_file(csv_writer, 'tmp.xlsx')
			os.remove('tmp.xlsx')

if __name__ == '__main__':
	main()
