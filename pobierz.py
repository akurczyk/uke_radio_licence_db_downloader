import urllib.request
import openpyxl
import csv
import re
import os

URL = 'https://www.uke.gov.pl/pozwolenia-radiowe-dla-klasycznych-sieci-radiokomunikacji-ruchomej-ladowej-5458'
WYRAZENIE = '<a href="/files/\?id_plik=([0-9]+)">[a-zA-Z0-9_-]+\.xlsx<\/a>'
PREFIX = 'https://www.uke.gov.pl/files/?id_plik='

def pobierz_liste_plikow():
	polaczenie = urllib.request.urlopen(URL)
	zestaw_znakow = polaczenie.info().get_content_charset()
	zawartosc_strony = polaczenie.read().decode(zestaw_znakow)
	polaczenie.close()

	lista_plikow = []
	for identyfikator_pliku in re.findall(WYRAZENIE, zawartosc_strony):
		lista_plikow.append(PREFIX + identyfikator_pliku)

	return lista_plikow

def zapisz_naglowki(zapis_do_csv):
	naglowki = []
	naglowki.append('Nr pozwolenia')
	naglowki.append('Data wygaśnięcia')
	naglowki.append('Nazwa stacji')
	naglowki.append('Rodzaj stacji')
	naglowki.append('Rodzaj sieci')
	naglowki.append('Długość geograficzna')
	naglowki.append('Szerokość geograficzna')
	naglowki.append('Promień obszaru obsługi')
	naglowki.append('Lokalizacja stacji')
	naglowki.append('ERP')
	naglowki.append('Azymut')
	naglowki.append('Elewacja')
	naglowki.append('Polaryzacja')
	naglowki.append('Zysk anteny')
	naglowki.append('Wysokość umieszczenia anteny')
	naglowki.append('Wysokość terenu')
	naglowki.append('Charakterystyka promieniowania - poziom')
	naglowki.append('Charakterystyka promieniowania - pion')
	naglowki.append('Częstotliwości nadawcze')
	naglowki.append('Częstotliwości odbiorcze')
	naglowki.append('Szerokości kanałów nadawczych')
	naglowki.append('Szerokości kanałów odbiorczych')
	naglowki.append('Operator')
	naglowki.append('Adres operatora')
	zapis_do_csv.writerow(naglowki)

def przetworz_plik(zapis_do_csv, nazwa_pliku):
	skoroszyt = openpyxl.load_workbook(nazwa_pliku)
	for arkusz in skoroszyt:
		pierwszy_rekord = True
		for rekord in arkusz:
			if not pierwszy_rekord:
				przetworz_rekord(zapis_do_csv, rekord)
			else:
				pierwszy_rekord = False

def przetworz_rekord(zapis_do_csv, rekord):
	rekord_do_csv = []
	for komorka in rekord:
		rekord_do_csv.append(komorka.value)
	rekord_do_csv[5] = zamien_wspolrzedne(rekord_do_csv[5])
	rekord_do_csv[6] = zamien_wspolrzedne(rekord_do_csv[6])
	zapis_do_csv.writerow(rekord_do_csv)

def zamien_wspolrzedne(wsp):
	if not wsp: return ''
	wsp = re.split('E|N|\'|"', wsp)
	return int(wsp[0]) + int(wsp[1])/60 + int(wsp[2])/3600

def main():
	print('Otwieranie pliku nadajniki.csv do zapisu...')
	plik_csv = open('nadajniki.csv', 'w', newline='', encoding='utf-8')
	zapis_do_csv = csv.writer(plik_csv)

	print('Dopisywanie wiersza nagłówkowego do pliku...')
	zapisz_naglowki(zapis_do_csv)

	for plik in pobierz_liste_plikow():
		print('Przetwarzanie pliku ' + plik + '...')
		urllib.request.urlretrieve(plik, 'tmp.xlsx')
		przetworz_plik(zapis_do_csv, 'tmp.xlsx')
		os.remove('tmp.xlsx')
	
	print('Zamykanie pliku nadajniki.csv...')
	plik_csv.close()

if __name__ == '__main__':
	main()