#!/usr/bin/python
" dump an entire database into one MS-Excel file with nice formatting "
# this works best if you used CREATE VIEW in the database to make
# pseudotables that end-users want to see
#
# Requirements;
# - `apt-get install python-mysqldb python-xlsxwriter`
# - or `pip install --user XlsxWriter`
# - your user & password entries in the [client] block of ~/.my.cnf

from math import log
import datetime
import MySQLdb
import os
import sys
import xlsxwriter


def _utf8ize_list(stuff):
  newstuff = []
  for i in stuff:
    if type(i) in (str, unicode):
      newstuff.append(i.encode('utf8'))
    else:
      newstuff.append(i)
  return newstuff


def _reguess_colwidths(old_widths, row):
  # assumes len(old_widths) == len(row)
  new_widths = list(old_widths)
  for i in range(len(new_widths)):
    if row[i] is None:
      length = 1
    elif isinstance(row[i], (str, unicode, bytearray, buffer)):
      length = len(row[i])
    elif isinstance(row[i], (int, long)):
      if row[i] == 0:
        length = 1
      else:
        length = log(row[i]) / log(10)
    elif isinstance(row[i], float):
      length = len(repr(row[i]))
    elif isinstance(row[i], datetime.datetime):
      length = 16
    elif isinstance(row[i], bool):
      length = 1
    else:
      print "WARN: unexpected data type %s %s" % (type(row[i]), repr(row[i]))
      length = len(repr(row[i]))
    if length > 216:
      # sanity check
      length = 216
    if length > new_widths[i]:
      new_widths[i] = length
  return new_widths


def db_init(db, user=None, passwd=None, host='localhost'):
  deffile = os.environ['HOME'] + '/.my.cnf'
  if user is None:
    dbh = MySQLdb.connect(
        host=host, db=db, use_unicode=True, charset='UTF8',
        read_default_file=deffile)
  else:
    dbh = MySQLdb.connect(
        host=host, user=user, passwd=passwd,
        db=db, use_unicode=True, charset='UTF8',
        read_default_file=deffile)
  return dbh


def zhu_li(dbname):
  " Does the thing "
  dbh = db_init(dbname)
  cur = dbh.cursor()
  cur.execute("""
      SELECT table_name from information_schema.tables
      WHERE table_catalog = 'def' and table_schema = '%s'
      ORDER BY table_name""" % dbname)
  tablenames = [i[0] for i in cur.fetchall()]
  try:
    os.unlink('%s.xlsx' % dbname)
  except:
    pass
  xlfile = xlsxwriter.Workbook('%s.xlsx' % dbname)
  xlsx_header_format = xlfile.add_format({'bold': True, 'bg_color': 'silver'})
  xlsx_float_format = xlfile.add_format({'num_format': '0.00'})
  xlsx_int_format = xlfile.add_format({'num_format': '0'})
  xlsx_date_format = xlfile.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
  # xlsx_text_format = xlfile.add_format({})
  for tablename in sorted(tablenames):
    cur.execute("SELECT * FROM %s" % tablename)
    colnames = [i[0] for i in cur.description]
    xlsheet = xlfile.add_worksheet(tablename[:31])
    # format columns
    for colnum in range(len(cur.description)):
      # can't set Autofit width, using default 'None'
      coltype = cur.description[colnum][1]
      if coltype in (MySQLdb.FIELD_TYPE.DOUBLE, MySQLdb.FIELD_TYPE.FLOAT):
        xlsheet.set_column(colnum, colnum, None, xlsx_float_format)
      if coltype in MySQLdb.NUMBER:
        xlsheet.set_column(colnum, colnum, None, xlsx_int_format)
      elif coltype in (MySQLdb.DATETIME, MySQLdb.DATE):
        xlsheet.set_column(colnum, colnum, None, xlsx_date_format)
      else:
        # xlsheet.set_column(colnum, colnum, None, xlsx_text_format)
        pass
    # format header line
    xlsheet.set_row(0, cell_format=xlsx_header_format)
    # keep row 0 pinned to the top
    xlsheet.freeze_panes(row=1, col=0)
    # write header row
    rownum = 0
    #xlsheet.write_row(row=rownum, col=0, data=_utf8ize_list(colnames))
    xlsheet.write_row(row=rownum, col=0, data=colnames)
    colwidths = [len(i) * .8 for i in colnames]
    # write the rest of the rows
    for row in cur.fetchall():
      rownum += 1
      #xlsheet.write_row(row=rownum, col=0, data=_utf8ize_list(row))
      xlsheet.write_row(row=rownum, col=0, data=row)
      colwidths = _reguess_colwidths(colwidths, row)
    for colnum in range(len(cur.description)):
      # default font is not monospaced, the ".8" is a kludge
      xlsheet.set_column(colnum, colnum, (colwidths[colnum] + 1) * .8)
  xlfile.close()
  print "Created %s.xlsx" % dbname


if (len(sys.argv) != 2):
  print "Usage: $0 dbname -- writes dump to dbname.xlsx"
  print "  assumes you already wrote your credentials to ~/.my.cnf"
  sys.exit(1)
else:
  zhu_li(sys.argv[1])  # "Zhu Li!  Do the thing!"  -- Varrick
