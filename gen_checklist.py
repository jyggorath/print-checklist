#!/usr/bin/python3
from sys import argv
from sys import stderr
from os import remove
from os.path import dirname
from os.path import realpath
from os.path import isfile
from json import load
from json import dumps
from datetime import datetime
from datetime import timedelta
from time import sleep
from argparse import ArgumentParser
from win32api import ShellExecute
from win32print import GetDefaultPrinter


class DummyObject:
	pass
def serialize(data: dict, serialize_dicts=False) -> DummyObject:
	object = DummyObject()
	for key, value in data.items():
		if type(value) is dict and serialize_dicts:
			setattr(object, key, serialize(value))
		else:
			setattr(object, key, value)
	return object

def gen_dates(start: datetime, end: datetime) -> list:
	dates = [start]
	d = start
	while d < end:
		d += timedelta(days=1)
		dates.append(d)
	return dates

def yesno(message: str) -> bool:
	a = None
	while a != 'Y' and a != 'n':
		a = input('%s (Y/n): ' % message)
	return (a == 'Y')

def scanchars_to_rtf(string: str) -> str:
	# "æøå ÆÅØ": \'e6\'f8\'e5 \'c6\'d8\'c5
	s = string.replace('æ', "\\'e6")
	s = s.replace('ø', "\\'f8")
	s = s.replace('å', "\\'e5")
	s = s.replace('Æ', "\\'c6")
	s = s.replace('Ø', "\\'d8")
	s = s.replace('Å', "\\'c5")
	s = s.replace('Ã¦', "\\'e6")	# æ (win/PS)
	s = s.replace('Ã¸', "\\'f8")	# ø (win/PS)
	s = s.replace('Ã¥', "\\'e5")	# å (win/PS)
	s = s.replace('Ã†', "\\'c6")	# Æ (win/PS)
	s = s.replace('Ã˜', "\\'d8")	# Ø (win/PS)
	s = s.replace('Ã…', "\\'c5")	# Å (win/PS)
	return s

def gen_heading(config: DummyObject, date: datetime, rtf=False) -> str:
	day = config.translations[date.strftime('%A')]
	month = config.translations[date.strftime('%B')]
	heading = config.heading.replace('DAY', day).replace('MONTH', month).replace('DATE', str(date.day))
	if rtf:
		heading = scanchars_to_rtf(heading)
	return heading


parser = ArgumentParser(usage=' %s  [--config-file|-c CONFIG_FILE]  [--start-date|-s STARTDATE]  [--end-date|-e ENDDATE]  [--out-file|-o OUT_FILE]  [--help|-h]  [--dump-config|-d]  [-n|--no-print]' % argv[0])
parser.add_argument(
	'-c', '--config-file', 
	default='%s/config.json' % dirname(realpath(argv[0])), 
	help='config file (defaults to %s/config.json' % dirname(realpath(argv[0]))
)
parser.add_argument(
	'-s', '--start-date', 
	default='tomorrow', 
	help='start date, format: YYYY-MM-DD (defaults to the next day)'
)
parser.add_argument(
	'-e', '--end-date', 
	default='weekend', 
	help='end date, format: YYYY-MM-DD (defaults to first sunday, or next sunday, if start date is first sunday)'
)
parser.add_argument(
	'-o', '--out-file', 
	default='checklist.rtf', 
	help='filename of the .rtf file to generate (default: checklist.rtf)'
)
parser.add_argument(
	'-d', '--dump-config', 
	action='store_true', 
	help='dumps the current config'
)
parser.add_argument(
	'-n', '--no-print', 
	action='store_true', 
	help="generate the document, but don't print it"
)
arguments = parser.parse_args()


if not isfile(arguments.config_file):
	print('Config file not found', file=stderr)
	exit(1)
with open(arguments.config_file, 'r') as fh:
	config = serialize(load(fh))


if arguments.dump_config:
	print(dumps(config.__dict__, indent='\t', ensure_ascii=False))
	exit()


if arguments.start_date == 'tomorrow':
	start = datetime.now()
	start += timedelta(days=1)
else:
	try:
		start = datetime.strptime(arguments.start_date, '%Y-%m-%d')
	except ValueError as exception:
		print('Invalid start date:', exception.args[0], file=stderr)
		exit(1)


if arguments.end_date == 'weekend':
	end = start
	if end.isoweekday() == 7:
		end += timedelta(days=1)
	while end.isoweekday() != 7:
		end += timedelta(days=1)
else:
	try:
		end = datetime.strptime(arguments.end_date, '%Y-%m-%d')
	except ValueError as exception:
		print('Invalid end date:', exception.args[0], file=stderr)
		exit(1)


if start > end:
	print('Invalid dates: End date must be later than start date', file=stderr)
	exit(1)


dates = gen_dates(start, end)

for date in dates:
	heading = gen_heading(config, date)
	print(heading)
	for task in config.daily:
		print('☐ ', task)
	if config.specific[str(date.isoweekday())] is not None:
		for task in config.specific[str(date.isoweekday())]:
			print('☐ ', task)
	print('')

if not yesno('Looks good?'):
	exit()

# https://www.oreilly.com/library/view/rtf-pocket-guide/9781449302047/ch01.html
rtf = '{\\rtf1\\ansi\deff0 {\\fonttbl {\\f0 Arial;}}\n'
for date in dates:
	heading = gen_heading(config, date, rtf=True)
	rtf += '\n\\f0\\fs56 {\\pard ' + heading + '\\par}\n'
	for i in range(len(config.daily)):
		task = scanchars_to_rtf(config.daily[i])
		rtf += '\\fs40 {\\pard \par}\n'
		rtf += "\\fs40 {\\pard {\\rtlch\\fcs1 \\af44 \\ltrch\\fcs0 \\f44\\insrsid13374959\\charrsid13374959 \\u9744\\'3f}  " + task
		if i == (len(config.daily) - 1) and config.specific[str(date.isoweekday())] is None and date != end:
			rtf += '\\page\par}\n'
		else:
			rtf += '\\par}\n'
	if config.specific[str(date.isoweekday())] is not None:
		for i in range(len(config.specific[str(date.isoweekday())])):
			task = scanchars_to_rtf(config.specific[str(date.isoweekday())][i])
			rtf += '\\fs40 {\\pard \par}\n'
			rtf += "\\fs40 {\\pard {\\rtlch\\fcs1 \\af44 \\ltrch\\fcs0 \\f44\\insrsid13374959\\charrsid13374959 \\u9744\\'3f}  " + task
			if i == (len(config.specific[str(date.isoweekday())]) - 1) and date != end:
				rtf += '\\page\par}\n'
			else:
				rtf += '\\par}\n'
rtf += '}'

with open(arguments.out_file, 'w') as fh:
	fh.write(rtf)

print('Saved to', arguments.out_file)

if not arguments.no_print:
	print('Printing...')
	# http://timgolden.me.uk/python/win32_how_do_i/print.html
	ShellExecute(0, 'print', arguments.out_file, '/d:"%s"' % GetDefaultPrinter(), '.', 0)
	sleep(5)
	remove(arguments.out_file)
