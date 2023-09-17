from pylatex import Document, LongTable, MultiColumn, Figure, Package, NoEscape, UnsafeCommand, \
        LineBreak, MultiRow, PageStyle, Head, MiniPage, StandAloneGraphic, LargeText, NewPage
from pylatex.base_classes import CommandBase, Arguments, Command
from pylatex.utils import bold
from pylatex.basic import NewLine
from pylatex.position import VerticalSpace

from math import floor, ceil
import sys
print (sys.argv)
import os.path

import pandas as pd

numPlayersPerRow = 3
#picDir="./pics/"
picDir=""

class ExampleCommand(CommandBase):
    """
    A class representing a custom LaTeX command.

    This class represents a custom LaTeX command named
    ``exampleCommand``.
    """

    _latex_name = 'exampleCommand'
    packages = [Package('color')]

class PlayerInfo(CommandBase):
    _latex_name = 'playerInfo'

class PlayerCard(CommandBase):
    _latex_name = 'playercard'

def clear_nan(cell):
    if pd.isna(cell):
        return ""
    else :
        return str(cell)

def genenerate_longtabu():
    geometry_options = {
        "head": "60pt",
        "margin": "0.5in",
        "top": "2in",
#        "bottom": "0.6in",
#        "document_options": "12pt",
        "includeheadfoot": False
    }
    doc = Document(page_numbers=False, geometry_options=geometry_options)
#    doc = Document(documentclass='extarticle',page_numbers=False, geometry_options=geometry_options)
    doc.packages.append(Package('graphicx'))
    doc.packages.append(Package('color'))
    new_comm = UnsafeCommand('newcommand', '\playerInfo', options=3,
        extra_arguments=r'\color{red} #1 \newline \color{black} #2 \newline #3 \color{black}')
    doc.append(new_comm)
    new_comm = UnsafeCommand('newcommand', '\exampleCommand', options=3,
        extra_arguments=r'\color{#1} #2 #3 \color{black}')
    doc.append(new_comm)

#    doc.packages.append(Package('extsizes'))


#    doc.preamble.append(Command('usepackage', 'helvet'))
    doc.preamble.append(Command('usepackage', 'bookman'))
#    doc.preamble.append(Package('helvet'))
#    doc.preamble.append(Command('usepackage', 'palatino'))

    fileName=os.path.basename(sys.argv[1])
    outputName=str(os.path.splitext(str(fileName))[0])
    df=pd.read_excel(sys.argv[1],engine='openpyxl')
    print(df)

    totPlayers = len(df.index);
    wholeRows = floor(totPlayers/numPlayersPerRow)
    lastRowPlayers = totPlayers % numPlayersPerRow

    first_page = PageStyle("firstpage")

    with first_page.create(Head("L")) as header_left:
        with header_left.create(MiniPage(width=NoEscape(r"0.49\textwidth"),
                                         pos='c')) as logo_wrapper:
            #logo_file = os.path.join(os.path.dirname(__file__),sys.argv[4])
            logo_file = sys.argv[4]
            logo_wrapper.append(StandAloneGraphic(image_options="width=120px",
                                filename=logo_file))
    with first_page.create(Head("R")) as header_right:
        with header_right.create(MiniPage(width=NoEscape(r"0.49\textwidth"),
                                         pos='c',align='r')) as logo_wrapper:
            logo_file = os.path.join(os.path.dirname(__file__),
                                     '../pics/EGRL_WithWords_bluebackground.png')
            logo_wrapper.append(StandAloneGraphic(image_options="width=120px",
                                filename=logo_file))
    # Add document title
    with first_page.create(Head("C")) as center_header:
        with center_header.create(MiniPage(width=NoEscape(r"0.49\textwidth"),
                                 pos='c', align='c')) as title_wrapper:
            title_wrapper.append(LargeText(bold(str(sys.argv[2]))))
            title_wrapper.append(LineBreak())
            title_wrapper.append("    ")
            title_wrapper.append(LineBreak())
            #title_wrapper.append(LargeText(bold("New York 7's Roster")))
            #title_wrapper.append(LargeText(bold("Morris 7's Roster (Sept 17 2023)")))
            title_wrapper.append(LargeText(bold(str(sys.argv[3]))))
            title_wrapper.append(LineBreak())

    doc.preamble.append(first_page)
    doc.change_document_style("firstpage")
    doc.add_color(name="lightgray", model="gray", description="0.80")
    
    # https://jeltef.github.io/PyLaTeX/v1.2.0/examples/complex_report.html
    # https://stackoverflow.com/questions/65254535/xlrd-biffh-xlrderror-excel-xlsx-file-not-supported
    doc.append(VerticalSpace("1in"))
    #doc.append(LineBreak())


    # Generate data table
    with doc.create(LongTable("l l l l l l ")) as data_table:
            for i in range(wholeRows):
                   tmp1=str(df['Position'][3*i]) + "  " + str(df['First'][i*3])
                   tmp2=str(df['Position'][3*i+1]) + "  " + str(df['First'][i*3+1])
                   tmp3=str(df['Position'][3*i+2]) + "  " + str(df['First'][i*3+2])
                   pic1="'\includegraphics[width=1in]{" + picDir + str(df['pic'][i*3]) + "}"
                   pic2="'\includegraphics[width=1in]{" + picDir + str(df['pic'][i*3+1]) + "}"
                   pic3="'\includegraphics[width=1in]{" + picDir + str(df['pic'][i*3+2]) + "}"
                   data_table.add_row(MultiRow(4, data=NoEscape(pic1)),bold(tmp1), \
                        MultiRow(4, data=NoEscape(pic2)),bold(tmp2), \
                        MultiRow(4, data=NoEscape(pic3)),bold(tmp3), \
                        )
                   data_table.add_row('',bold(df['Last'][i*3]),'',bold(df['Last'][i*3+1]),'',bold(df['Last'][i*3+2]))
                   data_table.add_row('',clear_nan(df['Grade'][i*3]),'',clear_nan(df['Grade'][i*3+1]),'',clear_nan(df['Grade'][i*3+2]))
                   data_table.add_row('',clear_nan(df['Officer'][i*3]),'',clear_nan(df['Officer'][i*3+1]),'',clear_nan(df['Officer'][i*3+2]))
                   data_table.add_row('',clear_nan(df['Fifth'][i*3]),'',clear_nan(df['Fifth'][i*3+1]),'',clear_nan(df['Fifth'][i*3+2]))
                   data_table.add_row('','','','','','')
                   data_table.add_row('','','','','','')
                   data_table.add_row('','','','','','')
            i = i + 1

            if lastRowPlayers == 2 :
                   tmp1=str(df['Position'][3*i]) + "  " + str(df['First'][i*3])
                   tmp2=str(df['Position'][3*i+1]) + "  " + str(df['First'][i*3+1])
                   pic1="'\includegraphics[width=1in]{" + picDir + str(df['pic'][i*3]) + "}"
                   pic2="'\includegraphics[width=1in]{" + picDir + str(df['pic'][i*3+1]) + "}"
                   data_table.add_row(MultiRow(4, data=NoEscape(pic1)),bold(tmp1), \
                        MultiRow(4, data=NoEscape(pic2)),bold(tmp2), \
                        '','', \
                        )
                   data_table.add_row('',bold(df['Last'][i*3]),'',bold(df['Last'][i*3+1]),'','')
                   data_table.add_row('',clear_nan(df['Grade'][i*3]),'',clear_nan(df['Grade'][i*3+1]),'','')
                   data_table.add_row('',clear_nan(df['Officer'][i*3]),'',clear_nan(df['Officer'][i*3+1]),'','')
                   data_table.add_row('',clear_nan(df['Fifth'][i*3]),'',clear_nan(df['Fifth'][i*3+1]),'','')

            elif lastRowPlayers == 1 :
                   tmp1=str(df['Position'][3*i]) + "  " + str(df['First'][i*3])
                   pic1="'\includegraphics[width=1in]{" + picDir + str(df['pic'][i*3]) + "}"
                   data_table.add_row(MultiRow(4, data=NoEscape(pic1)),bold(tmp1), \
                        '','', \
                        '','', \
                        )
                   data_table.add_row('',bold(df['Last'][i*3]),'','','','')
                   data_table.add_row('',clear_nan(df['Grade'][i*3]),'','','','')
                   data_table.add_row('',clear_nan(df['Officer'][i*3]),'','','','')
                   data_table.add_row('',clear_nan(df['Fifth'][i*3]),'','','','')
      
                        

    # https://tex.stackexchange.com/questions/533471/pylatex-create-custom-command
    # https://www.jason-french.com/blog/2012/01/17/using-figures-within-tables-in-latex/
    doc.append(NewPage())
    
    doc.generate_pdf(outputName, clean_tex=False)

genenerate_longtabu()
