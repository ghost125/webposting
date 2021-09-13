import streamlit as st
from multiapp import MultiApp
from apps import createspec, updatexml # import your app modules here


app = MultiApp()

# Add all your application here
app.add_app("Create Specification", createspec.app)
app.add_app("Update the XML", updatexml.app)


# The main app
app.run()