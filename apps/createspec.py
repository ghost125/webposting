import base64
import os
from io import BytesIO
import datetime

from .xml2spec import xmlparse,createSpec
import streamlit as st 
from docx import Document

def specdownload(docxfile):
    buffered = BytesIO()
    file = open(docxfile,'rb').read()
    buffered.write(file)
    b64 = base64.b64encode(buffered.getvalue()).decode()
    href = f'<a href="data:file/docx;base64,{b64}" download="spec.docx">Download Docx file</a>'
    return href


def app():
    st.title('Create SPEC')
    st.write('This page is used to create the Spec base on the XML file from web posting team')
    inxml = st.file_uploader("Upload XML file",key='beunqiue')
    if st.button('Create spec',key='run'):
        test = xmlparse(inxml)
        spec = createSpec()
        overllage=test.get_population_agegroup()
        country=test.get_country()
        spec.add_trial_information(overllage,country)
        disp,group_disp=test.get_disposition()
        spec.add_subject_disposition(group_disp,disp)
        group_baseline,set_baseline,baseline=test.get_baseline2()
        spec.add_baseline(group_baseline,set_baseline,baseline)
        group_endp=test.get_outcome_groups()
        endp,endp_a=test.get_outcome()
        sets = test.get_outcome_sets()
        spec.add_endpoints(endp,group_endp,sets,endp_a)
        filename = 'sp'+datetime.datetime.now().strftime("%Y%m%d%h%m%c").replace(':','').replace(' ','')+'.docx'
        spec.Save(filename)
        st.markdown(specdownload(filename),unsafe_allow_html=True)
        os.remove(filename)

