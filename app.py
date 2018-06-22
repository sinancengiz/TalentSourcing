# Module used to connect Python with MongoDb
import pymongo
# Dependencies
import pandas as pd
import os
import json
from flask import Flask, render_template
# The default port used by MongoDB is 27017
# https://docs.mongodb.com/manual/reference/default-mongodb-port/
conn = 'mongodb://localhost:27017'
client = pymongo.MongoClient(conn)

# Store filepath in a variable
data_file = os.path.join("data.xlsx")

# Read our Data file with the pandas library
# Not every CSV requires an encoding, but be aware this can come up
data_file_df = pd.read_excel(data_file, encoding="ISO-8859-1")

records = json.loads(data_file_df.to_json(orient='records'))

# Define the 'talentSourceDB' database in Mongo
db = client.talentSourceDB

# Drops collection if available to remove duplicates
db.talentsTable.drop()

#Inset data to TalentsTable
db.talentsTable.insert_many(records)

talent_data = db.talentsTable.find()

# Create an instance of our Flask app.
app = Flask(__name__)

# Set route
@app.route('/')
def index():
    # Store the entire talent_datam collection in a list
    talentslist = list(talent_data)
   
    # Return the template with the talent_data list passed in
    return render_template('index.html', talentslist=talentslist)


if __name__ == "__main__":
    app.run(debug=True)