# generate email template


# define dictionaries
collections = {
    "general":
    "On behalf of Prof. Samir Jaber, I'm delighted to invite you to contribute to Intensive Care Medicine.|",

    "brain":
    """On behalf of Pofs. Samir Jaber, Chiara Robba & Geert Meyfroidt, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming collection titled \"The Brain in the Line of Fire: Neuroprotection and Critical Illness\"|On behalf of<br>Prof. Kiara Robba, Deputy Editor and Collection Editor<br>Prof. Geert Meyfroidt, Collection Guest Editor<br>Prof. Samir Jaber, ICM Editor-in-Chief""",

    "standard care":
    """On behalf of Profs. Samir Jaber and Elie Azoulay, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming \"Standard of Care\" collection focusing on different areas of critical care, providing comprehensive insights into the current best practices and guidelines that define the standard of care in the field.|On behalf of<br>Prof. Elie Azoulay, Collection Guest Editor<br>Prof. Samir Jaber, ICM Editor-in-Chief""",

    "physiology": """On behalf of Profs. Samir Jaber, Julie Helms & Michael Darmon, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming collection titled \"Physiology in Critical Illness\".|On behalf of<br>Prof. Julie Helms, Section Editor & Collection Editor<br>Prof. Michael Darmon, Section Editor & Collection Editor<br>Prof. Samir Jaber, ICM Editor-in-Chief""",

    "perioperative":
    """On behalf of Profs. Samir Jaber, Chiara Robba, Stefan Schaller & Audrey de Jong, I'm delighted to invite  you to contribute to Intensive Care Medicine's upcoming collection titled \"Perioperative Care in the Intensive Care Unit\".|On behalf of<br>Prof. Chiara Robba, Deputy Director & Collection Editor<br>Prof. Stefan Schaller, Section Editor & Collection Editor<br>Prof. Audrey de Jong, Collection Guest Editor<br>Prof. Samir Jaber, ICM Editor-in-Chief""",

    "obesity":
    """On behalf of Profs. Samir Jaber, Carol Hodgson, Gonzalo Hernandez & Emma Ridley, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming collection titled \"Obesity in the Intensive Care Unit\".|On behalf of<br>Prof. Carol Hodgson, Section Editor & Collection Editor<br>Prof. Gonzalo Hernandez, Section Editor & Collection Editor<br>Prof. Emma Ridley, Collection Guest Editor<br>Prof. Samir Jaber, ICM Editor-in-Chief"""
}

article_types = {
    "short editorial":
    """Key guidelines for the submission:<br>
    • Max 1000 words and 15 references (recent sources preferred)<br>
    • 1 mandatory illustration (table or figure)<br>
    • 3 maximum authors from diverse centers and geographic zones<br>
    • Electronic supplementary materials are unlimited""",

    "review article":
    """Key guidelines for the submission:<br>
    • Max 4,000 words and 75 references (recent sources preferred)<br>
    • Abstract: unstructured (narrative) or structured (systematic reviews)<br>
    • 4-6 keywords; up to 5 illustrations (i.e., 3 figures & 2 tables)<br>
    • Between 12 and 15 authors, from diverse centers and geographic zones<br>
    • Electronic supplementary materials are unlimited"""
}


# function that creates letter salutation from a list of authors
def salute(authors):
    salutation = ''
    counter = 0
    if len(authors) <= 3:
        for author in authors:
            if len(author.split()) == 1:
                surname = author.split()[0]
            else:
                surname = ' '.join(author.split()[1:])
            if counter < len(authors)-1:
                salutation += "Dear Dr " + surname + ',<br>'
            else:
                salutation += "Dear Dr " + surname + ','
            counter += 1
    else:
        salutation += 'Dear All,'
    return salutation.strip() + '<br>'


# function capitalizing first letter of each word in a
# sentence but leaves all other letters untouched
def first_letter_to_cap(cad):
    words = cad.split()
    phrase = ''
    for word in words:
        phrase += word[0].upper() + word[1:] + ' '
    return phrase.rstrip()


def create_list_of_authors(task):
    authors = [task.author1.strip()]
    if task.author2 != None and task.author2.replace(' ', '') != '':
        authors.append(task.author2.strip())
    if task.author3 != None and task.author3.replace(' ', '') != '':
        authors.append(task.author3.strip())
    return authors


def get_signature():
    with open('signature.txt', 'r', encoding='utf-8') as f:
        html_string = f.read()
    return html_string


def generate_email_text(task):
    msg = ''
    article_type = task.type.lower()
    article_title = first_letter_to_cap(task.title.strip())
    date_format = '%d/%m/%Y'
    deadline = task.deadline.strftime(date_format)
    # determine the task collection according to name
    cad = set(task.collection.lower().strip().split())
    if {'standard', 'care'} <= cad:
        task_collection = 'standard care'
    elif {'brain'} <= cad:
        task_collection = 'brain'
    elif {'physiology'} <= cad:
        task_collection = 'physiology'
    elif {'perioperative'} <= cad:
        task_collection = 'perioperative'
    elif {'obesity'} <= cad:
        task_collection = 'obesity'
    else:
        task_collection = 'general'

    # create list of authors
    authors = create_list_of_authors(task)
    # append salute to authors
    msg += salute(authors) + '<br>'
    # append first paragraph
    msg += collections[task_collection].split('|')[0] + ' '
    # append text common to all emails
    cad = "Your expertise is highly regarded and we would greatly value a " + article_type + \
        " from you on \"" + article_title + "\". " + \
        "You may refine the title as you see fit.<br><br>"
    msg += cad
    # append submission guidelines
    msg += article_types[article_type] + "<br><br>"
    # append text with deadline
    msg += "Submission would be due by " + deadline + \
        ", but we can accommodate for some flexibility.<br><br>"
    msg += "Your expertise would make a significant impact, and we hope we'll have the privilege to read you in ICM.<br>"
    msg += "Kindly confirm your participation within seven days, and always feel free to reach out to the Editorial office for any questions or requests.<br><br>"
    msg += "Thank you for considering this invitation.<br><br>Kind regards,<br>ICM Editorial Office<br><br>"
    # append signatories
    msg += collections[task_collection].split('|')[1]
    msg += get_signature() + '<br>'

    return msg


def generate_FOLLOW_UP_email_text(task):
    msg = ''
    # get info from the task
    article_type = task.type.lower()
    article_title = first_letter_to_cap(task.title.strip())
    authors = create_list_of_authors(task)
    # create email text
    msg += salute(authors) + '<br>' + "I hope you are doing well. I'm writing to kindly follow up on our earlier invitation to contribute to the " + article_type + " on " + article_title + \
        " for ICM.<br>Would you still be interested to draft such a piece? Of course, the Editorial Office would be happy to assist you in inviting your co-authors.<br><br>Feel free to reach out to us for any question!<br><br>Kind regards,<br>Martin"

    msg += get_signature() + '<br>'

    return msg
