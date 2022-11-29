from fpdf import FPDF

pdf = FPDF()
pdf.add_page()

# Lane Header #

# Lane Footer #

# Written Part of narrative #
#   -Bring in using a manually generated txt file with written narrative sections
#   -Check for existing txt file in current folder, if it doesnt exist create a new file template

# Lists showing started/completed activities #
#   -Generate from activity csv export

# Tables showing relationship and activity changes #
#   -Build a heading for all categories and create a null result
#   -Pull data from schedule comparison export csv



pdf.set_font('arial','B',16)
pdf.cell(40,10,'Hello World!')
pdf.output('NarrativeTest.pdf','F')


