import argparse
import json
import xlsxwriter as x
import logging

def main():
	ap = argparse.ArgumentParser()
	ap.add_argument('infile', help='JSON formatted file')
	ap.add_argument('outfile', help='Path to put the csv file')
	ap.add_argument('-v', '--verbose', help='Make output verbose',
					dest='verbose', action='store_const',
					default=logging.WARNING, const=logging.INFO)
	ap.add_argument('-l', '--labels', help='Don\'t resolve labels to their name',
					action='store_false', default=True, dest='labels')
	ap.add_argument('-i', '--info', help='Add list info to sheet',
					action='store_true', default=False, dest='info')
	ap.add_argument('--add-empty', default=False, action='store_true',
					help='Add empty lists to sheet.', dest='add_empty')
	args = ap.parse_args()

	logging.basicConfig(level=args.verbose)

	json_sheet = None
	logging.info("Loading JSON file")
	with open(args.infile) as fp:
		json_sheet = json.load(fp)
	logging.info("File loaded and serialized")

	# Resolve labels
	if args.labels:
		json_sheet = resolve_labels(json_sheet)

	# Transform the cards entry into a more useuful format where cards are
	# keyed by list ID rather than in order. Keeps us from repeatedly iterating
	# through the whole thing
	tidied_json = get_cards(json_sheet, args)

	# Put stuff into worksheets
	# 1 sheet per list
	logging.info("Populating workbook")
	wb = x.Workbook(args.outfile)
	format_header = wb.add_format({'bold':True})
	for l in tidied_json["lists"]:
		if l["id"] not in tidied_json["cards"].keys() and not args.add_empty:
			logging.info("Ignoring empty list {}.".format(l["name"]))
			continue
		r = 0
		# Add the sheet, named by the list
		logging.info("Adding sheet {}".format(l["name"]))
		sheet = wb.add_worksheet(l["name"])
		# Add a row with all the info for the list
		if args.info:
			r = add_list_info(sheet, l, format_header, r)
		try:
			# Add a header row
			r = add_header(sheet, tidied_json["cards"][l["id"]], format_header, r)
		except KeyError:
			# This is an empty list
			continue
		# Add all the cards, based on the list ID
		add_cards(sheet, tidied_json["cards"][l["id"]], r)

	logging.info("Writing workbook")
	# write the file
	wb.close()
	print("File written to {}".format(args.outfile))

def resolve_labels(sheet):
	logging.info("Building dict of labels")
	new_labels = {}
	for label in sheet["labels"]:
		new_labels[label["id"]] = label["name"]
	logging.info("Applying new labels.")
	sheet["labels"] = new_labels
	return sheet

def get_cards(sheet, args):
	logging.info("Finessing cards list to a more useful dict")
	new_cards = {}
	for card in sheet["cards"]:
		# Deal with labels first
		if args.labels:
			new_labels = []
			for label in card["labels"]:
				new_labels.append(sheet["labels"][label["id"]])
			if len(new_labels) is 1:
				new_labels = new_labels[0]
			elif len(new_labels) is 0:
				new_labels = None
			card["labels"] = new_labels
		# Change all list and dict types into strings
		for k in card.keys():
			if type(card[k]) in (list, dict):
				card[k] = str(card[k])
		# Check to see if we've seen a card from this list
		# If not, add both the list as the key and the card
		if not card["idList"] in new_cards.keys():
			new_cards[card["idList"]] = [card]
		# If so, add just the card
		else:
			new_cards[card["idList"]].append(card)
	logging.info("Done. Returning better cards.")
	sheet["cards"] = new_cards
	return sheet

def add_list_info(sheet, l, format_header, r):
	logging.info("Adding list info to sheet.")
	sheet.write_row(r, 0, list(l.keys()), format_header)
	sheet.write_row(r + 1, 0, list(l.values()))
	r += 2
	logging.info("List info added.")
	return r

def add_header(sheet, cards, format_header, r):
	logging.info("Adding header.")
	# Grab the first card and get its keys
	sheet.write_row(r, 0, list(cards[0].keys()), format_header)
	r += 1
	logging.info("Header added.")
	return r

def add_cards(sheet, cards, r):
	logging.info("Adding cards.")
	for card in cards:
		sheet.write_row(r, 0, list(card.values()))
		r += 1
	logging.info("All cards added.")

if __name__ == '__main__':
	main()