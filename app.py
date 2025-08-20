from flask import Flask, render_template, request, redirect, send_file, url_for
from flask_sqlalchemy import SQLAlchemy
from email_template import generate_email_text
from flask_mail import Mail, Message
from datetime import datetime
from io import BytesIO
import pandas as pd


# app setup
app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///database.db"
# create database
db = SQLAlchemy(app)
# mail settings
app.config['MAIL_SERVER'] = 'smtp.mail.ru'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USE_SSL'] = True
app.config['MAIL_USERNAME'] = 'correo.x@mail.ru'
app.config['MAIL_PASSWORD'] = 'SpB7bMxqqJdcd7YKSQIh'
mail = Mail(app)


# create MODEL for each row in database
class Task(db.Model):
    # assign values to fields
    id = db.Column(db.Integer, primary_key=True)
    journal = db.Column(db.String(100), nullable=False)
    collection = db.Column(db.String(50))
    type = db.Column(db.String(50))
    title = db.Column(db.String(120))
    deadline = db.Column(db.Date)
    author1 = db.Column(db.String(30))
    email1 = db.Column(db.String(30))
    author2 = db.Column(db.String(30))
    email2 = db.Column(db.String(30))
    author3 = db.Column(db.String(30))
    email3 = db.Column(db.String(30))
    date_invited = db.Column(db.Date)
    status = db.Column(db.String(30))
    last_change_in_notes = db.Column(db.Date)
    notes = db.Column(db.String(2500))

    def __repr__(self) -> str:
        return f"Article Request {self.id}"


# define function that turns excel file into dataframe
def turn_excel_file_into_df_for_grid(excel_file):
    '''turns an excel file into a suitably
    formatted df for the table grid'''
    # define columns for final grid
    right_columns = ['journal', 'collection', 'type', 'title', 'deadline',
                     'author1', 'email1', 'author2', 'email2', 'author3',
                     'email3', 'date_invited', 'status', 'last_change_in_notes'
                     ]
    # define necessary dataframes
    df = pd.read_excel(excel_file)
    new_df = pd.DataFrame()
    # change original df column names if needed
    for c_title in df.columns:
        for r_col in right_columns:
            if r_col in c_title.replace(' ', '').lower():
                df.rename(columns={c_title: r_col}, inplace=True)
    # create new dataframe with all necessary columns
    for col in right_columns:
        # add existing columns
        if col in df.columns:
            # if column contains date, convert to python datetime format
            if col in {'deadline', 'date_invited', 'last_change_in_notes'}:
                new_df[col] = df[col].apply(lambda x: pd.Timestamp(x))
                new_df[col] = new_df[col].apply(
                    lambda x: x.to_pydatetime().date())
            # add a non-date column as it is
            new_df[col] = df[col]
        # create & add non-existing date column containing dummy date
        elif col in {'deadline', 'date_invited', 'last_change_in_notes'}:
            new_df[col] = pd.Timestamp('2000-01-01')
        # add a column containing an empty string
        else:
            new_df[col] = ''
    # return final dataframe
    return new_df


# define function to import excel file to database
def import_excel_to_db(excel_file):
    # get dataframe from excel file
    df = turn_excel_file_into_df_for_grid(excel_file)
    # insert dataframe rows as tasks
    for i in range(len(df)):
        # convert pandas timestamps to python datetime format
        deadline_py = df.iloc[i]['deadline'].to_pydatetime().date()
        date_invited_py = df.iloc[i]['date_invited'].to_pydatetime().date()
        last_change_in_notes_py = df.iloc[i]['last_change_in_notes'].to_pydatetime(
        ).date()
        # Create a new record
        try:
            task = Task(
                journal=df.iloc[i]['journal'],  # type: ignore
                collection=df.iloc[i]['collection'],  # type: ignore
                type=df.iloc[i]['type'],  # type: ignore
                title=df.iloc[i]['title'],  # type: ignore
                deadline=deadline_py,  # type: ignore
                author1=df.iloc[i]['author1'],  # type: ignore
                email1=df.iloc[i]['email1'],  # type: ignore
                author2=df.iloc[i]['author2'],  # type: ignore
                email2=df.iloc[i]['email2'],  # type: ignore
                author3=df.iloc[i]['author3'],  # type: ignore
                email3=df.iloc[i]['email3'],  # type: ignore
                date_invited=date_invited_py,  # type: ignore
                status=df.iloc[i]['status'],  # type: ignore
                last_change_in_notes=last_change_in_notes_py  # type: ignore
            )

            db.session.add(task)
            print(f'task {i} added to database commit')  # debugging

        except Exception as e:
            print(f"Error adding to commit => task {i}: {e}")
            db.session.rollback()
    try:
        db.session.commit()
        return True
    except Exception as e:
        db.session.rollback()
        print(f"Commit to database failed: {e}")
        return False


# route to main page (/)
@app.route('/', methods=['POST', 'GET'])  # type: ignore
def index():
    # add add an csv file with tasks
    if request.method == 'POST':
        my_file = request.files.get('file')
        if not my_file or my_file.filename == '':
            return ('No file selected!', 400)

        if my_file.filename.endswith('.xlsx'):  # type: ignore
            import_excel_to_db(my_file)
            return redirect('/')

    # see all current tasks
    else:
        tasks = Task.query.all()
        return render_template('index.html', tasks=tasks)


# route to confirm deletion
@app.route('/confirm-delete/<int:id>')
def confirm_delete(id):
    task = Task.query.get_or_404(id)
    return render_template('confirm_delete.html', task=task)


# route to delete an entry (row) in the data grid (database)
@app.route('/delete/<int:id>', methods=['POST'])
def delete_article(id):
    Task.query.filter_by(id=id).delete()
    db.session.commit()
    return redirect('/')


# route to export file in excel format
@app.route('/export-excel')
def export_excel():
    try:
        # Get data from database
        tasks = Task.query.all()
        # Convert to DataFrame (add your exact column names)
        data = [{
            'Journal': task.journal,
            'Collection': task.collection,
            'Article Type': task.type,
            'Article Title': task.title,
            'Deadline': task.deadline,
            'Author 1': task.author1,
            'Email 1': task.email1,
            'Author 2': task.author2,
            'Email 2': task.email2,
            'Author 3': task.author3,
            'Email 3': task.email3,
            'Date Invited': task.date_invited,
            'Status': task.status,
            'Last Change in Notes': task.last_change_in_notes,
            'Notes': task.notes
            # Add all other fields you want to export
        } for task in tasks]

        df = pd.DataFrame(data)

        # Create in-memory Excel file
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)  # Rewind the buffer

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name='tasks.xlsx',
            as_attachment=True
        )

    except Exception as e:
        return f"Export failed: {str(e)}", 500


# route to edit email template
@app.route('/edit-email/<int:task_id>', methods=['GET', 'POST'])
def edit_email(task_id):
    task = Task.query.get_or_404(task_id)
    # generate email content
    email_content = generate_email_text(task)

    if request.method == 'POST':
        # send the email
        final_content = request.form['email_content']
        msg = Message(
            subject=f"ICM Article Request {task.title}",
            sender='correo.x@mail.ru',
            recipients=[task.email1],
            body=final_content
        )
        mail.send(msg)
        flash('Email sent successfully!')  # type: ignore
        return redirect(url_for('index'))

    return render_template('edit_email.html',
                           email_content=email_content,
                           task=task)


# route to edit task
@app.route('/edit-task/<int:task_id>', methods=['GET', 'POST'])
def edit_task(task_id):
    # get task from database
    task = Task.query.get_or_404(task_id)
    current_year = datetime.now().year

    if request.method == 'POST':
        # update regular fields
        task.journal = request.form['journal']
        task.collection = request.form['collection']
        task.title = request.form['title']
        task.type = request.form['type']
        task.author1 = request.form['author1']
        task.email1 = request.form['email1']
        task.author2 = request.form['author2']
        task.email2 = request.form['email2']
        task.author3 = request.form['author3']
        task.email3 = request.form['email3']
        task.status = request.form['new_status']

        # update date_invited
        task.date_invited = datetime(
            year=int(request.form['date_invited_year']),
            month=int(request.form['date_invited_month']),
            day=int(request.form['date_invited_day'])
        ).date()

        # update deadline (if field exists)
        if all(f'deadline_{part}' in request.form for part in ['year', 'month', 'day']):
            task.deadline = datetime(
                year=int(request.form['deadline_year']),
                month=int(request.form['deadline_month']),
                day=int(request.form['deadline_day'])
            ).date()

        # update task notes
        notes = request.form['notes'][:2500]
        task.notes = notes
        task.last_change_in_notes = datetime.now().date()

        # commit to database
        db.session.commit()
        return redirect(url_for('index'))

    # Prepare date ranges for dropdowns
    date_ranges = {
        'days': list(range(1, 32)),
        'months': list(range(1, 13)),
        'invited_years': list(range(current_year - 5, current_year + 1)),
        'deadline_years': list(range(current_year, current_year + 3))
    }

    # Set default dates if None
    default_date = datetime.now().date()
    date_invited = task.date_invited or default_date
    deadline = task.deadline or default_date

    return render_template('edit_task.html', task=task, date_invited=date_invited, deadline=deadline, **date_ranges)


# runner and debugger
if __name__ == '__main__':

    with app.app_context():
        db.create_all()

    app.run(debug=True)
