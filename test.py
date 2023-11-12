import pandas as pd

filename_raw = "TM30 10-11-23-ข้อมูลดิบ.xls"

filename = "TM30 10-11-23.xls"

df = pd.read_excel(filename_raw)

# Selecting and renaming columns
df = df[['First Name\nชื่อ', 
         'Middle Name\nชื่อกลาง', 
         'Last Name\nนามสกุล',
         'Sex\nเพศ\n(M-ชาย,F-หญิง)',
         'Passport No.\nหนังสือเดินทาง',
         'Nationality\nสัญชาติ',
         'Date of Birth (mm/dd/yyyy) (A.D.)\nวันที่เกิด (ค.ศ.)',
         'Expire Date of Stay (mm/dd/yyyy) (A.D.)\nวันครบกำหนดอนุญาต (ค.ศ.)',
         ]]

# Renaming columns
df = df.rename(columns={
    'First Name\nชื่อ': 'ชื่อ\nFirst Name *',
    'Middle Name\nชื่อกลาง': 'ชื่อกลาง\nMiddle Name',
    'Last Name\nนามสกุล': 'นามสกุล\nLast Name',
    'Sex\nเพศ\n(M-ชาย,F-หญิง)': 'เพศ\nGender *',
    'Passport No.\nหนังสือเดินทาง': 'เลขหนังสือเดินทาง\nPassport No. *',
    'Nationality\nสัญชาติ': 'สัญชาติ\nNationality *',
    'Date of Birth (mm/dd/yyyy) (A.D.)\nวันที่เกิด (ค.ศ.)': 'วัน เดือน ปี เกิด\nBirth Date *\nDD/MM/YYYY(ค.ศ. / A.D.)',
    'Expire Date of Stay (mm/dd/yyyy) (A.D.)\nวันครบกำหนดอนุญาต (ค.ศ.)': 'วันที่แจ้งออกจากที่พัก\nCheck-out Date *\nDD/MM/YYYY(ค.ศ. / A.D.)',
})

# Add a new column 'เบอร์โทรศัพท์\nPhone No.' with NaN values
df['เบอร์โทรศัพท์\nPhone No.'] = ''

df.to_excel("test.xlsx", index=False, sheet_name="แบบแจ้งที่พัก Inform Accom")