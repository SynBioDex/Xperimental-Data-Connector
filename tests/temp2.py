from sre_constants import SUCCESS
# import xperimental_data_conv.main as xdc
# from xperimental-data-conv.xperimental_data_conv import main as xdc
from xperimental_data_converter.main import XDC
import os


fj_url = "charmmefj.synbiohub.org"
fj_user = ""
fj_pass = ""

sbh_url = "https://synbiohub.org"
sbh_user = "test@test.test"
sbh_pass = "test123"
sbh_collec = "XDC_package_test"

test_file_path ='../test_files'
excel_path = os.path.join(test_file_path, 'Medias.xlsm')

homespace = 'https://synbiohub.org/synbiotest'

fj_overwrite = False
sbh_overwrite=False

xdc = XDC(input_excel_path = excel_path,
            fj_url = fj_url,
            fj_user = fj_user, 
            fj_pass = fj_pass, 
            sbh_url = sbh_url, 
            sbh_user = sbh_user, 
            sbh_pass = sbh_pass, 
            sbh_collection = sbh_collec, 
            sbh_collection_description = 'XDC package test collection',
            sbh_overwrite = sbh_overwrite, 
            fj_overwrite = fj_overwrite, 
            homespace = homespace,
            fj_token = None, 
            sbh_token = None)

class Test_XDC(unittest.TestCase):
    def test_initialize(self):
        xdc.initialize()


  
      
        

