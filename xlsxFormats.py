import openpyxl as op
import openpyxl.styles as opStyles

def getFormats():
    # load all the base formats for fonts, borders etc 
    # then load the styles to use on cells
    return setupStyles(setupBaseFormats())  


def setupStyles(formats):
    # These are the actual styles to use in the code
    # they use fonts borders and number_formats etc as created below
    #
    # formats['header] is a dict with font and border styling  you can use on a cell
    #
    # formats['fonts'] is a set of font styles that can be used in real styles
    formats.update({
        'header': {'font': formats['fonts']['header'], 'border': formats['borders']['thinBottom']},
        'rowHeader': {'font': formats['fonts']['header'], 'border': formats['borders']['thinRight']},
        'memberHeader':  {'font': formats['fonts']['memberHeader'], 'border': formats['borders']['thinBottomBlue']},
        'memberName': {'font': formats['fonts']['memberName']},
        'memberValue':  {'font': formats['fonts']['memberValue']},
        'boolean1': {'font': formats['fonts']['boolean1']},
        'boolean0': {'font': formats['fonts']['boolean0']},
        'percentage': {'number_format': formats['number_formats']['percentage']},
    })
    return formats


def setupBaseFormats():
    # The common fonts borders and number_formats are set up here
    return {
        'fonts': {
            'header': opStyles.Font(color="444444", size=14, bold=True),
            'memberHeader': opStyles.Font(color="333388", size=12, bold=True),
            'memberName': opStyles.Font(color="333388", size=10, bold=True),
            'memberValue': opStyles.Font(color="000000", name='Courier',  size=10, bold=False),
            'boolean1': opStyles.Font(color="000000", name='Courier', size=10, bold=True),
            'boolean0': opStyles.Font(color="888888", name='Courier', size=10, bold=False)
        },
        'borders': {'thinBox': op.styles.Border(left=op.styles.Side(style='thin'),
                                                right=op.styles.Side(style='thin'),
                                                top=op.styles.Side(style='thin'),
                                                bottom=op.styles.Side(style='thin')),
                    'thickBottom': op.styles.Border(bottom=op.styles.Side(style='thick', color='000000')),
                    'thinBottom': op.styles.Border(bottom=op.styles.Side(style='thin', color='000000')),
                    'thinRight': op.styles.Border(right=op.styles.Side(style='thin', color='000000')),
                    'thinBottomBlue': op.styles.Border(bottom=op.styles.Side(style='thin', color='000088'))
                    },
        'number_formats': {'percentage': '0.00%'}
    }



