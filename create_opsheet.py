import pandas as pd
from fpdf import FPDF
from tkinter import filedialog, Tk

PG_WDTH = 188

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 350)
pd.set_option('max_colwidth', 40)

class PDF(FPDF):
    # TODO: Add Work order to header (and footer)
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

def gen_mat_dict(df):
    """Generates a dict of required materials and their amounts

    Args:
        df (dataframe): op specific dataframe
    Returns:
        mat_dict (Dict): Material ID: (len, width)
    """
    # Generates a dict of required materials and their amounts
    # returns Material ID : (len, width)
    mat_dict = {}
    for i, row in df.iterrows():
        # saw width is 1.0 in workorder, default to 0 if saw
        len = row['Length'] * row['Qty']
        wid = row['Width'] * row['Qty']
        old_tup = mat_dict.get(row['Material NO'], (0.0,0.0))
        mat_dict[row['Material NO']] = (old_tup[0] + len, old_tup[1] + wid)
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
    #for mat, dims in sorted(mat_dict.items()):
    #    mat_desc = df[df['Material NO'] == mat]['Material Description'].head(1).to_string(index=False)
    #    uom = df[df['Material NO'] == mat]['UOM'].head(1).to_string(index=False)
    #    pdf.cell(PG_WDTH/6, 10, f"{mat}", border = 1, ln = 0)
    #    pdf.cell(PG_WDTH/2, 10, f"{mat_desc}", border = 1, ln = 0)
    #    pdf.cell(PG_WDTH/3, 10, f"{dims[0]:.2f}x{dims[1]:.2f} ({dims[0]*dims[1]:.2f}) {uom}", border = 1, ln = 1)
    #pdf.ln(10)

    # Op table
    for mat in sorted(mat_dict.keys()):
        mat_desc = df[df['Material NO'] == mat]['Material Description'].head(1).to_string(index=False)
        pdf.set_font('Times', 'B', 12)
        pdf.ln(5)
        pdf.cell(0, 10, f"{san(mat)} : {san(mat_desc)}", ln=1)
        pdf.set_font('Times', '', 12)

        mat_df = df.loc[df["Material NO"] == mat].sort_values(by=['ID'])
        for i, row in mat_df.iterrows():
            # TODO: include part dimensions
            uom = row['UOM']
            next_op = row['op2'] if not pd.isna(row['op2']) else 'N/A'
            pdf.cell(PG_WDTH/8, 21, f"{san(row['Qty'])}x  ", border = 1, ln = 0, align='R')
            pdf.multi_cell(PG_WDTH-PG_WDTH/8, 7, f"{san(row['ID'])} : {san(row['Description'])}\n"\
                           f"Dimensions: {row['Length']:.2f} x {row['Width']:.2f} "\
                           f"({row['Length']*row['Width']:.2f}) {uom}\n"\
                           f"Next Op: {next_op}", border = 1)


if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

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

    # saw
    saw_df = df[df['op1'] == 'Saw']
    saw_mat_dict = gen_mat_dict(saw_df)
    # laser
    laser_df = df[df['op1'] == 'Laser']
    laser_mat_dict = gen_mat_dict(laser_df)
    # waterjet
    wj_df = df[df['op1'] == 'Sub-Water']
    wj_mat_dict = gen_mat_dict(wj_df)

    pdf = PDF()
    pdf.bom_number = file_path.split('/')[-1].split('.')[0] 

    pdf.alias_nb_pages()

    # material table at top
    write_op_pdf(saw_df, saw_mat_dict, "Saw", pdf)
    write_op_pdf(laser_df, laser_mat_dict, "Laser", pdf)
    write_op_pdf(wj_df, wj_mat_dict, "Water Jet", pdf)

    pdf.output(f"/users/brendan/Downloads/{pdf.bom_number}-opsheet.pdf", 'F')




