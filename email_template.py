# generate email template for martin's app
# TODO : format the date to display month in letters


# define dictionaries
collections = {
    "general":
    "On behalf of Prof. Samir Jaber, I'm delighted to invite you to contribute to Intensive Care Medicine.|",

    "brain":
    """On behalf of Pofs. Samir Jaber, Chiara Robba & Geert Meyfroidt, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming collection titled \"The Brain in the Line of Fire: Neuroprotection and Critical Illness\"|On behalf of\nProf. Kiara Robba, Deputy Editor and Collection Editor\nProf. Geert Meyfroidt, Collection Guest Editor\nProf. Samir Jaber, ICM Editor-in-Chief""",

    "standard care":
    """On behalf of Profs. Samir Jaber and Elie Azoulay, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming \"Standard of Care\" collection focusing on different areas of critical care, providing comprehensive insights into the current best practices and guidelines that define the standard of care in the field.|On behalf of\nProf. Elie Azoulay, Collection Guest Editor\nProf. Samir Jaber, ICM Editor-in-Chief""",

    "physiology": """On behalf of Profs. Samir Jaber, Julie Helms & Michael Darmon, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming collection titled \"Physiology in Critical Illness\".|On behalf of\nProf. Julie Helms, Section Editor & Collection Editor\nProf. Michael Darmon, Section Editor & Collection Editor\nProf. Samir Jaber, ICM Editor-in-Chief""",

    "perioperative":
    """On behalf of Profs. Samir Jaber, Chiara Robba, Stefan Schaller & Audrey de Jong, I'm delighted to invite  you to contribute to Intensive Care Medicine's upcoming collection titled \"Perioperative Care in the Intensive Care Unit\".|On behalf of\nProf. Chiara Robba, Deputy Director & Collection Editor\nProf. Stefan Schaller, Section Editor & Collection Editor\nProf. Audrey de Jong, Collection Guest Editor\nProf. Samir Jaber, ICM Editor-in-Chief""",

    "obesity":
    """On behalf of Profs. Samir Jaber, Carol Hodgson, Gonzalo Hernandez & Emma Ridley, I'm delighted to invite you to contribute to Intensive Care Medicine's upcoming collection titled \"Obesity in the Intensive Care Unit\".|On behalf of\nProf. Carol Hodgson, Section Editor & Collection Editor\nProf. Gonzalo Hernández, Section Editor & Collection Editor\nProf. Emma Ridley, Collection Guest Editor\nProf. Samir Jaber, ICM Editor-in-Chief"""
}

article_types = {
    "short editorial":
    """Key guidelines for the submission:
    • Max 1000 words and 15 references (recent sources preferred)
    • 1 mandatory illustration (table or figure)
    • 3 maximum authors from diverse centers and geographic zones
    • Electronic supplementary materials are unlimited""",

    "review article":
    """Key guidelines for the submission:
    • Max 4,000 words and 75 references (recent sources preferred)
    • Abstract: unstructured (narrative) or structured (systematic reviews)
    • 4-6 keywords; up to 5 illustrations (i.e., 3 figures & 2 tables)
    • Between 12 and 15 authors, from diverse centers and geographic zones
    • Electronic supplementary materials are unlimited"""
}


# salute function, takes in a list of authors
def salute(authors):
    salutation = '\n'
    counter = 0
    if len(authors) <= 3:
        for author in authors:
            if len(author.split()) == 1:
                surname = author.split()[0]
            else:
                surname = ' '.join(author.split()[1:])
            if counter < len(authors)-1:
                salutation += "Dear Dr " + surname + ',\n'
            else:
                salutation += "Dear Dr " + surname + ','
            counter += 1
    else:
        salutation += 'Dear All,'
    return salutation


def generate_email_text(task):
    msg = ''
    article_type = task.type.lower()
    article_title = task.title.strip().title()
    # TODO: in deadline, month should be displayed with letters
    deadline = str(task.deadline)
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
    authors = [task.author1.strip()]
    if task.author2 != None and task.author2.replace(' ', '') != '':
        authors.append(task.author2.strip())
    if task.author3 != None and task.author3.replace(' ', '') != '':
        authors.append(task.author3.strip())

    # append salute to authors
    msg += salute(authors) + '\n\n'
    # append first paragraph
    msg += collections[task_collection].split('|')[0] + ' '
    # append text common to all emails
    cad = "Your expertise is highly regarded and we would greatly value a " + article_type + \
        " from you on \"" + article_title + "\". " + \
        "You may refine the title as you see fit.\n\n"
    msg += cad
    # append submission guidelines
    msg += article_types[article_type] + "\n"
    # append text with deadline
    msg += "Submission would be due by " + deadline + \
        ", but we can accommodate for some flexibility.\n\n"
    msg += "Your expertise would make a significant impact, and we hope we'll have the privilege to read you in ICM.\n"
    msg += "Kindly confirm your participation within seven days, and always feel free to reach out to the Editorial office for any questions or requests.\n\n"
    msg += "Thank you for considering this invitation.\n\nKind regards,\nICM Editorial Office\n\n"
    # append signatories
    msg += collections[task_collection].split('|')[1] + '\n\n'

    return msg
