import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

def generate_sample_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    sample_dir = os.path.join(base_dir, "sample_data")
    photos_dir = os.path.join(sample_dir, "photos")
    os.makedirs(photos_dir, exist_ok=True)
    
    # 21 sample staff records matching exact user structure
    staff_data = [
        {"PHOTO": "NEELAVENI", "NAME": "S.NEELAVENI", "RF ID NO": 277735, "STD": "PRIMARY ASSISTANT", "DOB": "11.06.1976", "FATHER": "N.Kannan", "MOBILE": 9176750670, "ADD1": "No. 22, Thulukkanathamman Koil 1st Lane,", "ADD2": "Agaram, Chennai - 600082", "BG": ""},
        {"PHOTO": "SHAHITHA BEGUM", "NAME": "SAHITHA BEGUM.S", "RF ID NO": 435495, "STD": "PRIMARY ASSISTANT", "DOB": "15.07.1978", "FATHER": "N.Lathif Ahamed", "MOBILE": 9345841353, "ADD1": "No. 24, Subramani Garden 5th Street,", "ADD2": "Perambur, Chennai - 11", "BG": "B+ve"},
        {"PHOTO": "MOHAMMED USHAINA", "NAME": "MOHAMMED USHAINA.A", "RF ID NO": 589210, "STD": "SECONDARY ASSISTANT", "DOB": "20.03.1982", "FATHER": "A.Mohammed", "MOBILE": 9840123456, "ADD1": "15/7 North Street,", "ADD2": "Royapuram, Chennai - 600013", "BG": "O+ve"},
        {"PHOTO": "ANNALAKSHMI", "NAME": "ANNALAKSHMI.M", "RF ID NO": 312450, "STD": "GRADUATE TEACHER", "DOB": "05.11.1985", "FATHER": "M.Murugan", "MOBILE": 9789012345, "ADD1": "No. 8 Gandhi Road,", "ADD2": "T.Nagar, Chennai - 600017", "BG": "A+ve"},
        {"PHOTO": "DEVISRI", "NAME": "DEVISRI.K", "RF ID NO": 419823, "STD": "PRIMARY ASSISTANT", "DOB": "12.09.1988", "FATHER": "K.Kumar", "MOBILE": 9444567890, "ADD1": "Plot 42, Green Avenue,", "ADD2": "Velachery, Chennai - 600042", "BG": "AB+ve"},
        {"PHOTO": "KAVITHA", "NAME": "KAVITHA.R", "RF ID NO": 623411, "STD": "PET TEACHER", "DOB": "03.04.1980", "FATHER": "R.Rajan", "MOBILE": 9150112233, "ADD1": "12 Church Lane,", "ADD2": "Mylapore, Chennai - 600004", "BG": "B-ve"},
        {"PHOTO": "RAMESH", "NAME": "RAMESH.V", "RF ID NO": 781290, "STD": "LAB ASSISTANT", "DOB": "18.12.1979", "FATHER": "V.Viswanathan", "MOBILE": 9884556677, "ADD1": "45 Lake View Road,", "ADD2": "Adyar, Chennai - 600020", "BG": "O-ve"},
        {"PHOTO": "SARAVANAN", "NAME": "SARAVANAN.P", "RF ID NO": 192834, "STD": "SECONDARY ASSISTANT", "DOB": "25.01.1983", "FATHER": "P.Perumal", "MOBILE": 9600112244, "ADD1": "89 Station Road,", "ADD2": "Tambaram, Chennai - 600045", "BG": "A-ve"},
        {"PHOTO": "MEENAKSHI", "NAME": "MEENAKSHI.S", "RF ID NO": 543210, "STD": "HEADMISTRESS", "DOB": "14.02.1972", "FATHER": "S.Subramanian", "MOBILE": 9443123456, "ADD1": "100 Trunk Road,", "ADD2": "Punamallee, Chennai - 600056", "BG": "O+ve"},
        {"PHOTO": "VENKATESH", "NAME": "VENKATESH.G", "RF ID NO": 678901, "STD": "OFFICE ASSISTANT", "DOB": "08.08.1987", "FATHER": "G.Ganesan", "MOBILE": 9710998877, "ADD1": "33 School Street,", "ADD2": "Guindy, Chennai - 600032", "BG": "B+ve"},
        {"PHOTO": "LATHA", "NAME": "LATHA.B", "RF ID NO": 234567, "STD": "GRADUATE TEACHER", "DOB": "22.10.1981", "FATHER": "B.Baskaran", "MOBILE": 9841887766, "ADD1": "67 Bazaar Street,", "ADD2": "Saidapet, Chennai - 600015", "BG": "A1B+ve"},
        {"PHOTO": "SURESH", "NAME": "SURESH.N", "RF ID NO": 890123, "STD": "PRIMARY ASSISTANT", "DOB": "19.06.1986", "FATHER": "N.Narayanan", "MOBILE": 9940332211, "ADD1": "5 Cross Street,", "ADD2": "Kodambakkam, Chennai - 600024", "BG": "O+ve"},
        {"PHOTO": "GAYATHRI", "NAME": "GAYATHRI.K", "RF ID NO": 345678, "STD": "SECONDARY ASSISTANT", "DOB": "31.05.1990", "FATHER": "K.Krishnan", "MOBILE": 9003114455, "ADD1": "78 Main Road,", "ADD2": "Chromepet, Chennai - 600044", "BG": "B+ve"},
        {"PHOTO": "JAYANTHI", "NAME": "JAYANTHI.T", "RF ID NO": 901234, "STD": "PRIMARY ASSISTANT", "DOB": "09.07.1984", "FATHER": "T.Thirumalai", "MOBILE": 9176554433, "ADD1": "14 Market Street,", "ADD2": "Aminjikarai, Chennai - 600029", "BG": "A+ve"},
        {"PHOTO": "BALAJI", "NAME": "BALAJI.D", "RF ID NO": 456789, "STD": "PET TEACHER", "DOB": "16.03.1983", "FATHER": "D.Dharmalingam", "MOBILE": 9840998811, "ADD1": "29 High Road,", "ADD2": "Egmore, Chennai - 600008", "BG": "O-ve"},
        {"PHOTO": "CHITRA", "NAME": "CHITRA.M", "RF ID NO": 112233, "STD": "GRADUATE TEACHER", "DOB": "27.09.1977", "FATHER": "M.Mani", "MOBILE": 9789223344, "ADD1": "52 Canal Bank Road,", "ADD2": "Mandaveli, Chennai - 600028", "BG": "B+ve"},
        {"PHOTO": "DINESH", "NAME": "DINESH.S", "RF ID NO": 445566, "STD": "LAB ASSISTANT", "DOB": "04.01.1989", "FATHER": "S.Sundaram", "MOBILE": 9444112233, "ADD1": "8 Cross Road,", "ADD2": "Ashok Nagar, Chennai - 600083", "BG": "A+ve"},
        {"PHOTO": "GOPAL", "NAME": "GOPAL.R", "RF ID NO": 778899, "STD": "OFFICE ASSISTANT", "DOB": "11.11.1985", "FATHER": "R.Ramachandran", "MOBILE": 9150998877, "ADD1": "91 Post Office Street,", "ADD2": "Vadapalani, Chennai - 600026", "BG": "O+ve"},
        {"PHOTO": "HEMA", "NAME": "HEMA.P", "RF ID NO": 998877, "STD": "PRIMARY ASSISTANT", "DOB": "28.08.1991", "FATHER": "P.Padmanabhan", "MOBILE": 9884112233, "ADD1": "17 New Street,", "ADD2": "West Mambalam, Chennai - 600033", "BG": "B+ve"},
        # Record 20 has missing photo on purpose
        {"PHOTO": "INDIRA_MISSING", "NAME": "INDIRA.K", "RF ID NO": 334455, "STD": "SECONDARY ASSISTANT", "DOB": "15.04.1986", "FATHER": "K.Kannan", "MOBILE": 9600998877, "ADD1": "63 Temple Avenue,", "ADD2": "Villivakkam, Chennai - 600049", "BG": "AB+ve"},
        # Record 21 has missing DOB & Mobile on purpose
        {"PHOTO": "JANAKI", "NAME": "JANAKI.S", "RF ID NO": 556677, "STD": "PRIMARY ASSISTANT", "DOB": "", "FATHER": "S.Srinivasan", "MOBILE": "", "ADD1": "41 Anna Street,", "ADD2": "Padi, Chennai - 600050", "BG": "O+ve"}
    ]
    
    # Save to Excel with sheet named 'STAFF'
    df = pd.DataFrame(staff_data)
    excel_path = os.path.join(sample_dir, "staff_data.xlsx")
    
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="STAFF", index=False)
        
    print(f"Sample Staff Excel generated at: {excel_path}")
    
    colors = [
        (41, 128, 185), (39, 174, 96), (142, 68, 173), (211, 84, 0),
        (192, 57, 43), (22, 160, 133), (44, 62, 80), (243, 156, 18),
        (52, 152, 219), (46, 204, 113)
    ]
    
    # Generate photos for records (skipping record index 19 INDIRA_MISSING)
    for i, staff in enumerate(staff_data):
        if i == 19: # Skip missing photo test case
            continue
            
        photo_name = str(staff["PHOTO"]) + ".jpg"
        filepath = os.path.join(photos_dir, photo_name)
        
        bg_color = colors[i % len(colors)]
        img = Image.new('RGB', (280, 350), color=bg_color)
        draw = ImageDraw.Draw(img)
        
        # Silhouette
        draw.ellipse([70, 60, 210, 200], fill=(255, 255, 255))
        draw.ellipse([25, 220, 255, 380], fill=(255, 255, 255))
        
        # Header banner
        draw.rectangle([15, 15, 265, 45], fill=(0, 0, 0))
        try:
            font = ImageFont.load_default()
        except Exception:
            font = None
            
        draw.text((30, 23), f"RF:{staff['RF ID NO']}", fill=(255, 255, 255))
        img.save(filepath)
        
    print(f"Generated sample photos in: {photos_dir}")

if __name__ == "__main__":
    generate_sample_data()
