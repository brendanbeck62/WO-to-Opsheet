import pandas as pd
import platform
from fpdf import FPDF
from tkinter import filedialog, Tk, messagebox
from os import getenv
import argparse 

PG_WDTH = 188

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 350)
pd.set_option('max_colwidth', 40)

class PDF(FPDF):
    op_title = ""
    bom_number = 0

    def header(self):
        self.set_font('Times', 'B', 12)
        self.cell(80, 15, f"bom: {self.bom_number}", 0, 0, 'C')
        self.set_font('Times', 'B', 15)
        self.cell(30, 10, self.op_title, 1, 0, 'C')
        self.ln(20)
        pass

    # Page footer
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

    def set_title(self, op_string):
        self.op_title = op_string

def get_args():
    parser = argparse.ArgumentParser(description="Convert Solidworks WOs to Opsheet PDFs")
    parser.add_argument('--debug', nargs='?', const=1, type=int, help="set debug level")
    parser.add_argument('-i', '--infile', help="optional in file for debugging")
    return parser.parse_args()

def gen_mat_dict(df):
    """Generates a dict of required materials and their amounts

    Args:
        df (dataframe): op specific dataframe
    Returns:
        mat_dict (Dict): Material ID: (len, width)
    """
    # Generates a dict of required materials and their amounts
    # returns Material ID : (len, width, sqft)
    mat_dict = {}
    for i, row in df.iterrows():
        # saw width is 1.0 in workorder, default to 0 if saw
        len = row['Length'] * row['Qty']
        wid = row['Width'] * row['Qty']

        sqft = row['Length'] * row['Width'] * row['Qty']

        old_tup = mat_dict.get(row['Material NO'], (0.0, 0.0, 0.0))
        mat_dict[row['Material NO']] = (old_tup[0] + len, old_tup[1] + wid, old_tup[2] + sqft)
        #print(f"{row['Level']}:{row['Material NO']}: ({old_tup[0] + len}, {old_tup[1] + wid}, {old_tup[2] + sqft})")
    return mat_dict

def san(string):
    string = str(string)
    return string.encode('latin-1', 'replace').decode('latin-1')

def write_op_pdf(df, mat_dict, op, pdf):

    pdf.set_title(op)
    pdf.add_page()

    #pdf.set_font('Times', 'B', 12)
    #pdf.cell(0, 10, "Materials:", ln=1)
    pdf.set_font('Times', '', 12)
    # material table
    for mat, dims in sorted(mat_dict.items()):
        mat_desc = df[df['Material NO'] == mat]['Material Description'].head(1).to_string(index=False)
        uom = df[df['Material NO'] == mat]['UOM'].head(1).to_string(index=False)
        if (uom == "FT"):
            pdf.cell(PG_WDTH/6, 10, f"{san(mat)}", border = 1, ln = 0)
            pdf.cell(PG_WDTH/2, 10, f"{san(mat_desc)}", border = 1, ln = 0)
            pdf.cell(PG_WDTH/3, 10, f"{dims[0]/12:.2f} {uom}", border = 1, ln = 1)
        else:
            # material table only requires square feet, which should be calculated by
            # summing the individual part's square feet, rather than summing the
            #length and width and then calculating the square feet
            pdf.cell(PG_WDTH/6, 10, f"{san(mat)}", border = 1, ln = 0)
            pdf.cell(PG_WDTH/2, 10, f"{san(mat_desc)}", border = 1, ln = 0)
            pdf.cell(PG_WDTH/3, 10, f"{dims[2]/144:.2f} {uom}", border = 1, ln = 1)

    # Op table
    for mat in sorted(mat_dict.keys()):
        mat_desc = df[df['Material NO'] == mat]['Material Description'].head(1).to_string(index=False)
        pdf.set_font('Times', 'B', 12)
        pdf.ln(5)
        pdf.cell(0, 10, f"{san(mat)} : {san(mat_desc)}", ln=1)
        pdf.set_font('Times', '', 12)

        mat_df = df.loc[df["Material NO"] == mat].sort_values(by=['ID'])
        for i, row in mat_df.iterrows():
            uom = row['UOM']
            next_op = row['op2'] if not pd.isna(row['op2']) else 'N/A'

            # Convert len and wid to feet, (right now values are in inches,
                # but labels are feet (SF / F)
            # This is because engineers work in inches, but stocking UOM is ft/sf.
            length = row['Length']/12
            width = row['Width']
            if uom != "FT":
                width = width/12

            pdf.cell(PG_WDTH/8, 21, f"{san(row['Qty'])}x  ", border = 1, ln = 0, align='R')

            # seperate logic for saw (feet) and everythign else (sq feet).
            if uom == "FT":
                pdf.multi_cell(PG_WDTH-PG_WDTH/8, 7, f"{san(row['ID'])} : {san(row['Description'])}\n"\
                               f"Dimensions per: {length:.2f} {uom}\n"\
                               f"Next Op: {next_op}", border = 1)
            else:
                pdf.multi_cell(PG_WDTH-PG_WDTH/8, 7, f"{san(row['ID'])} : {san(row['Description'])}\n"\
                               f"Dimensions per: {length:.2f} x {width:.2f} "\
                               f"({(length * width):.2f}) {uom}\n"\
                               f"Next Op: {next_op}", border = 1)


if __name__ == "__main__":

    args = get_args()

    root = Tk()
    root.withdraw()

    file_path = args.infile if args.infile else filedialog.askopenfilename()

    df = pd.DataFrame(pd.read_excel(file_path,
            converters={
                'ID': lambda x: str(x).strip(),
                'Qty': lambda x: int(x),
                'Length': lambda x: float(x),
                'Width': lambda x: float(x)
            },
            dtype = str,
        )
    )

    df = df[df['Part Type'] == 'Make']
    df = df.drop(columns=['Revision', 'Part Type', 'op4', 'op5', 'Sales Category', 'PROD Line', 'Price ID'])

    # update qty to include parent's qty (recursively)
    # because 6.2.10 has a qty of 3, but 6.2 has a qty of 2,
    # 6.2.10's effective qty is 6.

    # Additionally, the first row in the sheet is Level=0, which should act
    # as the parent for all integer (root level) parts.

    #Level       ID                  Desc                        Qty
    #0           1234-56-100         blah ASSEMBLY               1
    #6           9929-06-200         TROUGH ASSEMBLY             1
    #6.2         9929-06-300         TROUGH WELDMENT             2
    #6.2.1       650599-18-5040      5" SCH40 x 18" LG. 304SS    1
    #6.2.10      650598-18-304SS     18'' CROSS RIB, 304SS       3
    #6.2.11      9929-06-304-LH      STIFFENER ANGLE, LH         1
    for i, row in df.iterrows():
        # anything with a '.' is a child
        if '.' in row['Level']:
            level = row['Level']
            child_qty = row['Qty']
            # rfind gets index of last occurance
            parent_index = level.rfind('.')
            parent = level[:parent_index]
            parent_qty = df[df['Level'] == parent]['Qty']
            #print(f"parent = {df[df['Level'] == parent]}")
            #print(f"{level}, parent={parent}, parentqty={parent_qty}, childqty={child_qty}")
            df.at[i,'Qty'] = child_qty * parent_qty
        elif row['Level'] != '0':
            # I am a integer, so make my parent be 0
            level = row['Level']
            child_qty = row['Qty']
            parent_qty = df[df['Level'] == '0']['Qty']
            df.at[i,'Qty'] = child_qty * parent_qty

    # saw
    saw_df = df[df['op1'] == 'Saw']
    saw_mat_dict = gen_mat_dict(saw_df)
    # laser
    laser_df = df[df['op1'] == 'Laser']
    laser_mat_dict = gen_mat_dict(laser_df)
    # waterjet
    wj_df = df[df['op1'] == 'Waterjet']
    wj_mat_dict = gen_mat_dict(wj_df)
    # sub-water
    swj_df = df[df['op1'] == 'Sub-Water']
    swj_mat_dict = gen_mat_dict(swj_df)

    pdf = PDF()
    pdf.bom_number = file_path.split('/')[-1].split('.')[0]

    pdf.alias_nb_pages()

    write_op_pdf(saw_df, saw_mat_dict, "Saw", pdf)
    write_op_pdf(laser_df, laser_mat_dict, "Laser", pdf)
    write_op_pdf(wj_df, wj_mat_dict, "Water Jet", pdf)
    write_op_pdf(swj_df, swj_mat_dict, "Sub Water", pdf)

    # TODO: Save this to the same place the excel file was found, rather than the downloads
    if platform.system() == 'Windows':
        pdf.output(f"{getenv('USERPROFILE')}\\downloads\\{pdf.bom_number}-opsheet.pdf", 'F')
    else:
        pdf.output(f"/users/brendan/Downloads/{pdf.bom_number}-opsheet.pdf", 'F')

    messagebox.showinfo("Create Opsheet", f"Success!\nWrote pdf to \ndownloads/{pdf.bom_number}-opsheet.pdf")
