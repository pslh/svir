#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim: tabstop=4 shiftwidth=4 softtabstop=4
#
# Copyright (c) 2010-2012, GEM Foundation.
#
# OpenQuake is free software: you can redistribute it and/or modify it
# under the terms of the GNU Affero General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# OpenQuake is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with OpenQuake.  If not, see <http://www.gnu.org/licenses/>.

"""
Map full-column names to short, DB friendly names for SVIR data-set
"""
import sys

from xlrd import open_workbook

_SHEET_NAME = 'Database'
#
# Column IDs are in first column 'A' == index 0
# Column Names are in column 'D' == index 4
#
_COLUMN_ID_INDEX = 0
_COLUMN_NAME_INDEX = 4


def _load_data_sheet(filename, sheet_name):
    """
    Load the named excel file and sheet
    """
    sys.stderr.write('Loading {0}\n'.format(filename))
    excel = open_workbook(filename)
    sys.stderr.write('Done Loading, opening sheet {0}\n'.format(sheet_name))

    # obtain reference to "Database" sheet
    db_sheet = excel.sheet_by_name(sheet_name)
    sys.stderr.write('Done opening sheet {0}\n'.format(sheet_name))
    return db_sheet


def _output_sql_table_def(col_id, count):
    """
    Write out template for SQL table definition
    """
    if(count > 0):
        sys.stdout.write(',')
    sys.stdout.write('\n\t{0}\tDOUBLE PRECISION'.format(col_id))


def _output_atsv_header_names(col_id, count):
    """
    Write out CSV style header names using @ separator
    """
    if(count > 0):
        sys.stdout.write('@')
        sys.stdout.write('{0}'.format(col_id))


def _output_atsv_mapping_row(row, col_id, col_name):
    """
    Write out row number, column Id and column name as @ separated values

    """
    sys.stdout.write('{0}@{1}@{2}\n'.format(row, col_id, col_name))


def _output_mapping(db_sheet):
    """
    Loop over all rows in sheet outputing columnId and columnName
    """
    count = 0
    for row in range(2, db_sheet.nrows):
        col_id = db_sheet.cell_value(row, _COLUMN_ID_INDEX).encode('utf-8')
        col_name = db_sheet.cell_value(row, _COLUMN_NAME_INDEX).encode('utf-8')
        _output_atsv_mapping_row(row, col_id, col_name)
        # _output_atsv_header_names(col_id, count)
        # _output_sql_table_def(col_id,count)
        count = count + 1


def main():
    """
    Obtain filename from command line args, load sheet and output mapping
    """

    if(len(sys.argv) < 2):
        sys.stderr.write('Usage: {0} <excel filename>\n'.format(sys.argv[0]))
        sys.exit(-1)
    filename = sys.argv[1]

    # Load named excel file
    db_sheet = _load_data_sheet(filename, _SHEET_NAME)
    _output_mapping(db_sheet)

#
# Main driver
#
if __name__ == "__main__":
    main()
