from flask import Flask, render_template, request, redirect, send_file
# from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy
from flask_mail import Mail, Message
from datetime import datetime
from io import BytesIO
import pandas as pd


# TODO: add a function to create and clean the csv file;
# change date format to YY-MM-DD when creating the csv file;
# trim each entry from starting and ending blank spaces;


# app setup
app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///database.db"
# create database
db = SQLAlchemy(app)
app.config['MAIL_SERVER'] = 'smtp.mail.ru'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USE_SSL'] = True
app.config['MAIL_USERNAME'] = 'correo.x@mail.ru'
app.config['MAIL_PASSWORD'] = 'SpB7bMxqqJdcd7YKSQIh'
mail = Mail(app)


# create MODEL for each row in database
class ArticleRequest(db.Model):

    # assign values to fields
    id = db.Column(db.Integer, primary_key=True)
    journal = db.Column(db.String(100), nullable=False)
    collection = db.Column(db.String(50))
    type = db.Column(db.String(50))
    title = db.Column(db.String(120))
    author1 = db.Column(db.String(30))
    email1 = db.Column(db.String(30))
    author2 = db.Column(db.String(30))
    email2 = db.Column(db.String(30))
    author3 = db.Column(db.String(30))
    email3 = db.Column(db.String(30))
    date_invited = db.Column(db.Date)
    status = db.Column(db.String(30))
    last_change_in_notes = db.Column(db.Date)
    notes = db.Column(db.String(500))

    def __repr__(self) -> str:
        return f"Article Request {self.id}"


def parse_date(date_string):
    """Convert YYYY-MM-DD string to date object"""
    if date_string:  # Only parse if string is not empty
        return datetime.strptime(date_string, "%Y-%m-%d").date()
    return None  # Return None for empty strings


# define function to import csv data to data base
def import_csv_to_db(csv_data):
    # assume file contains no headers
    for row in csv_data:
        # Create a new record
        try:
            article_request = ArticleRequest(
                journal=row[0],  # type: ignore
                collection=row[1],  # type: ignore
                type=row[2],  # type: ignore
                title=row[3],  # type: ignore
                author1=row[4],  # type: ignore
                email1=row[5],  # type: ignore
                author2=row[6],  # type: ignore
                email2=row[7],  # type: ignore
                author3=row[8],  # type: ignore
                email3=row[9],  # type: ignore
                date_invited=parse_date(row[10]),  # type: ignore
                status=row[11],  # type: ignore
                last_change_in_notes=parse_date(row[12]))  # type: ignore

            db.session.add(article_request)
        except Exception as e:
            print(f"Error adding row {row}: {e}")
            db.session.rollback()
    try:
        db.session.commit()
        return True
    except Exception as e:
        db.session.rollback()
        print(f"Commit failed: {e}")
        return False


# route to main page
@app.route('/', methods=['POST', 'GET'])  # type: ignore
def index():
    # add add an csv file with article requests
    if request.method == 'POST':
        my_file = request.files.get('file')
        if not my_file or my_file.filename == '':
            return ('No file selected!', 400)
        if my_file.filename.endswith('.csv'):  # type: ignore
            # read csv file as a string & remove unnecesary quotation marks
            raw_csv_data = my_file.read().decode('utf-8').replace('"', '')
            # turn string into a list of lists (rows for the db)
            csv_data = [x.split(',') for x in raw_csv_data.split('\n')]
            import_csv_to_db(csv_data)
            return redirect('/')
    # see all current article requests
    else:
        art_requests = ArticleRequest.query.all()
        return render_template('index.html', article_requests=art_requests)


# route to confirm deletion
@app.route('/confirm-delete/<int:id>')
def confirm_delete(id):
    article_request = ArticleRequest.query.get_or_404(id)
    return render_template('confirm_delete.html', article_request=article_request)


# route to delete an entry (row) in the data grid (database)
@app.route('/delete/<int:id>', methods=['POST'])
def delete_article(id):
    ArticleRequest.query.filter_by(id=id).delete()
    db.session.commit()
    return redirect('/')


# route to change article request status
@app.route('/change-status/<int:id>', methods=['GET', 'POST'])
def change_status(id):
    article_request = ArticleRequest.query.get_or_404(id)

    if request.method == 'POST':
        # Update status
        article_request.status = request.form['new_status']
        db.session.commit()
        return redirect('/')

    return render_template('change_status.html', article_request=article_request)


# route to export file in excel format
@app.route('/export-excel')
def export_excel():
    try:
        # Get data from database
        article_requests = ArticleRequest.query.all()

        # Convert to DataFrame (add your exact column names)
        data = [{
            'Journal': article_request.journal,
            'Collection': article_request.collection,
            'Article Type': article_request.type,
            'Article Title': article_request.title,
            'Author 1': article_request.author1,
            'Email 1': article_request.email1,
            'Author 2': article_request.author2,
            'Email 2': article_request.email2,
            'Author 3': article_request.author3,
            'Email 3': article_request.email3,
            'Date Invited': article_request.date_invited,
            'Status': article_request.status,
            'Last Change in Notes': article_request.last_change_in_notes,
            'Notes': article_request.notes
            # Add all other fields you want to export
        } for article_request in article_requests]

        df = pd.DataFrame(data)

        # Create in-memory Excel file
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)  # Rewind the buffer

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name='article_requests.xlsx',
            as_attachment=True
        )

    except Exception as e:
        return f"Export failed: {str(e)}", 500


# route for notes editing
@app.route('/edit-notes/<int:id>', methods=['GET', 'POST'])
def edit_notes(id):
    article_request = ArticleRequest.query.get_or_404(id)

    if request.method == 'POST':
        notes = request.form['notes'][:500]
        article_request.notes = notes
        article_request.last_change_in_notes = datetime.now().date()
        db.session.commit()
        return redirect('/')

    return render_template('edit_notes.html', article_request=article_request)


# dictionary with email templates
EMAIL_TEMPLATES = {
    "template1": """Dear {title} {last_name},

We're about to publish the attached article "{original_title}" from Dr {original_author} and we'd like to invite you to draft a {article_type} on the topic of {article_title}.

Thank you,
Kind regards,
Martin""",

    "template2": """Hello {title} {last_name},

We're excited to share the article "{original_title}" by Dr {original_author} and would appreciate your input as a {article_type} contributor on "{article_title}".

Best wishes,
Martin"""
}


# route to select email template
@app.route('/draft-email/<int:article_request_id>')
def select_template(article_request_id):
    article_request = ArticleRequest.query.get_or_404(article_request_id)
    return render_template('select_template.html', article_request=article_request)


# route to edit email template
@app.route('/edit-email/<int:article_request_id>/<template_name>', methods=['GET', 'POST'])
def edit_email(article_request_id, template_name):
    article_request = ArticleRequest.query.get_or_404(article_request_id)

    # Prepare template variables
    template_vars = {
        'title': 'Dr.',
        'last_name': article_request.author1.split()[-1],
        'original_title': article_request.title,
        'original_author': article_request.author1,
        'article_type': article_request.type,
        'article_title': article_request.title
    }

    if request.method == 'POST':
        # Process the edited email
        msg = Message(
            'Martin\'s Article Request',
            sender='correo.x@mail.ru',
            recipients=[article_request.email1],
            body=request.form['email_content']
        )
        mail.send(msg)
        return redirect('/')
    # Get the selected template and fill in variables
    email_content = EMAIL_TEMPLATES[template_name].format(**template_vars)
    return render_template('edit_email.html',
                           email_content=email_content,
                           article_request=article_request)


# runner and debugger
if __name__ == '__main__':

    with app.app_context():
        db.create_all()

    app.run(debug=True)
