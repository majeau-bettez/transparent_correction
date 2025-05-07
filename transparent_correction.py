import os
import getpass
import logging  # For warning
from datetime import datetime
import time
import re
import pandas as pd
import numpy as np
import shutil  # Notably for copyfile
from pathlib import Path
import seaborn as sns
import matplotlib.pyplot as plt

import jenkspy  # Jenks clustering

import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

# Set / fix styles issues
sns.set_style()
pd.DataFrame._repr_latex_ = lambda self: """\centering{}""".format(self.to_latex())

import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

# Constants

_FAIL_LETTER = 'f'
_ALL_LETTERS = ['f', 'd', 'd+', 'c', 'c+', 'b', 'b+', 'a', 'a+']
_ALL_PASSING_LETTERS = ['d', 'd+', 'c', 'c+', 'b', 'b+', 'a', 'a+']
_ALL_PASSING_NORMAL_LETTERS = ['d', 'd+', 'c', 'c+', 'b', 'b+', 'a']

# Text strings for automated messages

_txt_salutation = """\
        <html>
          <body>
          <p> Bonjour {} {}, <br><br>
"""

_txt_end = """    </p>
                  </body>
              </html>
              """

_txt_score_overview = """

La moyenne du groupe est de {:.1f} points, et sa note médiane est de {:.1f} points. Vous avez obtenu une note de {:.1f} points. 
<hr>

Voici le détail de vos points:
""".replace('\n', '<br>')

_txt_score_details = """
    {} : {}
        {} points sur {}
    """.replace('\n', '<br>')

_txt_mistakes_details = """
<hr>

Et voici le détail des points perdus (toute pénalité négative correspond à un point 'gagné' ou 'bonus'):

{}

<hr> """.replace('\n', '<br>')


def correction_parser(filename, exam_name):
    """ Reads in correction template and generates `Grader` object

    Parameters
    ----------
    filename : str
    exam_name : str

    Returns
    -------
    Grader object

    """
    with pd.ExcelFile(filename) as f:
        raw = pd.read_excel(f, sheet_name='Corrections', header=0, index_col=0)
        raw_codes = pd.read_excel(f, sheet_name='codes', header=0, index_col=[0, 1])
        universal_codes = pd.read_excel(f, sheet_name='codes_universels', header=0, index_col=0)
        tmp = pd.read_excel(f, sheet_name='totaux', header=0, index_col=0)
        totals = tmp.loc[:, 'points']
        try:
            titles = tmp.loc[:, 'titre'].fillna('')
        except KeyError:
            titles = pd.Series('', index=totals.index, name='titre').fillna('')


        init = pd.read_excel(f, sheet_name='init', header=0, index_col=0).squeeze()
        versions = pd.read_excel(f, sheet_name='versions', header=0, index_col=0).squeeze()

    return Grader(exam_name, raw, raw_codes, universal_codes, totals, init, titles, versions)


class Grader:

    def __init__(self, exam_name, raw_corr, raw_codes, universal_codes, totals, init, titles, versions=None):
        """
        Grader class to contain raw correction data and processed grades

        Parameters
        ----------
        exam_name : str
            Name of exam, used for documentation
        raw_corr: DataFrame
            Correction codes for each student (row) and each question (columns) read from template
        raw_codes: DataFrame
            Definition and weight of each correction/error code; A negative "penalty" means that points are being added.
        universal_codes : DataFrame
            Definition and weight (relative and/or absolute penalty) of error codes that can apply to any question
        totals : Data Series
            Total points for each question
        versions: DataFrame (optional)
            For each question (row) and each version of the exam (columns, 'A', 'B', etc.) the name/number of the
            question as seen by the student on the exam.
        """
        # Data from tremplate
        self.exam_name = exam_name
        self.raw_corr = raw_corr
        self.raw_codes = raw_codes
        self.universal_codes = universal_codes
        self.totals = totals
        self.init = init
        self.titles = titles
        self.versions = versions

        # Constants - hardcoded
        self._OK_TERMS = ['ok', '0', '', 'OK']
        self._CONTACT_COLS = ['prénom', 'nom', 'courriel']
        self._COLS_TO_ALWAYS_DROP = ['version', 'Exemple/explication']

        # Refined data
        contacts = raw_corr[self._CONTACT_COLS]
        corr = raw_corr.drop(self._CONTACT_COLS, axis=1)

        # Remove any student not present
        absent = corr.isna().all(1)
        logging.warning("Removing these student ID's, as clearly absent from exam: {}".format(
            list(absent[absent].index)))
        self.corr = corr.loc[~absent]
        self.contacts = contacts.loc[~absent]
        maybe_absent = self.corr.loc[:, self.totals > 0].isna().any(1)
        if maybe_absent.any():
            logging.warning("The following ID's have NaN in their correction, absent from exam?: {}".format(
                list(maybe_absent[maybe_absent].index)))


        # Semi-final variables
        self.correction_matrix = None
        self.codes = None

        # Variables finales
        self.grades = None

        self.message = dict()
        self.message['salutation'] = _txt_salutation
        self.message['foreword'] = ""
        self.message['score_overview'] = _txt_score_overview
        self.message['score_details'] = _txt_score_details
        self.message['mistakes_details'] = _txt_mistakes_details
        self.message['closing'] = ""

    def calc_grades(self, cols_to_drop=None):
        """
        The main method. Calculates the grade of each student.

        Calls in sequence `_pivot_corr()`, `_clean_codes()`, `_check_sanity_and_harmonize()`, and `_calc_grades()`

        Parameters
        ----------
        cols_to_drop : list
            Columns in the template to ignore in the calculation process (custom additionnal columns, etc.)

        """
        if cols_to_drop is None:
            cols_to_drop = []

        self._pivot_corr(cols_to_drop=cols_to_drop)
        self._clean_codes(cols_to_drop=cols_to_drop)
        self._check_sanity_and_harmonize()
        self._calc_grades()

    def _pivot_corr(self, cols_to_drop=None):
        """ Pivot the correction comments in raw_corr into a binary matrix

         For each student, we go from a list of error codes to a binary matrix indicating which students (rows) did
          what mistake (columns, level 1) in what question (columns, level 0)

        Parameters
        ----------
        cols_to_drop : list
            Columns in the template to ignore in the calculation process (custom additionnal columns, etc.)
        """

        cols_to_drop = self._COLS_TO_ALWAYS_DROP + cols_to_drop
        # Define dataframe with multi-index columns capturing all correction types
        for cx in self.corr.columns:
            all_codes = self.corr[cx].dropna().unique().tolist()
            all_codes = {y for x in all_codes for y in _clean(x)}
            new_cx = pd.MultiIndex.from_product([[cx], all_codes])
            try:
                mcx = mcx.append(new_cx)
            except NameError:
                mcx = new_cx

        correction_matrix = pd.DataFrame(0.0, index=self.corr.index, columns=mcx)

        # Fill
        for ix in self.corr.index:
            for cx in self.corr.columns:
                erreurs = _clean(self.corr.loc[ix, cx])
                for err in erreurs:
                    correction_matrix.loc(axis=0)[ix].loc[[cx], err] += 1

        # Remove other stuff
        if cols_to_drop is not None:
            correction_matrix = correction_matrix.drop(cols_to_drop, axis=1, level=0, errors='ignore')

        correction_matrix = correction_matrix.reindex(columns=self.totals.index, level=0)

        # Enlève OK
        self.correction_matrix = correction_matrix.drop(labels=self._OK_TERMS, axis=1, level=1, errors='ignore')

    def _clean_codes(self, cols_to_drop):
        """ Clean correction codes and weighting (drop empty, expand universal correction codes, etc.)

        Parameters
        ----------
        cols_to_drop : list
            Columns in the template to ignore in the calculation process (custom additionnal columns, etc.)
        """

        # Drop extraneous columns
        cols_to_drop = self._COLS_TO_ALWAYS_DROP + cols_to_drop
        codes = self.raw_codes.drop(cols_to_drop, axis=1, errors='ignore')

        # Drop empty rows
        todrop = codes['points'].isna()
        self.codes = codes.loc[~todrop]

        # Scale and insert universal codes
        for ix, v in self.totals.items():
            scaled_codes = (self.universal_codes[['pénalités relatives', 'pénalités_absolues']] * [v, 1])
            selected_codes = scaled_codes.dropna(axis=0, how='all').min(axis=1, skipna=True).to_frame('points')
            new_codes = pd.concat([self.universal_codes['définition'], selected_codes], axis=1, join='inner')
            new_codes.index = pd.MultiIndex.from_product([[ix], new_codes.index])
            self.codes = pd.concat([self.codes, new_codes])
        

        duplicate_codes = self.codes.index.duplicated()
        if duplicate_codes.any():
            duplicates = self.codes.index[duplicate_codes]
            raise ValueError(f"There are duplicate error code definitions: {duplicates}.")
        else:
            self.codes.reindex(index=self.totals.index, level=0)

    def distribute_codes(self, common_ix, applicables):
        """ Expand code list if a common correction code is applicable to more than one question (but not all).

        Parameters
        ----------
        common_ix : str
            A common correction code already defined in self.codes for a question
        applicables: list
            List of other questions for which this code is also applicable
        """
        common_codes = self.codes.loc[common_ix]
        tmp = pd.concat([common_codes, ] * len(applicables), keys=applicables)
        codes = pd.concat([self.codes, tmp], axis=0).sort_index()
        self.codes = codes.drop(common_ix, axis=0)

    def _check_sanity_and_harmonize(self):
        """ Inspect for missing correction codes & harmonize dimensions. """
        missing_codes = self.correction_matrix.columns.difference(self.codes.index)
        if len(missing_codes):
            logging.warning("Certains codes de correction ne sont pas définis: {}".format(missing_codes))
        else:
            self.codes = self.codes.reindex(self.correction_matrix.columns)

        # Check that all questions are appropriately represented
        qx1 = set(self.codes.index.levels[0])
        qx2 = set(self.correction_matrix.columns.levels[0])
        qx3 = set(self.totals.index)
        qx4 = set(self.init.index)

        if qx1 == qx2 == qx3 == qx4:
            pass
        else:
            logging.warning("La liste des questions n'est pas la même pour tous les onglets.")


    def _calc_grades(self):
        """ Calculate grades: the init - the sum of the weighted corrections, with minimum of 0 (no negative grades)"""
        penalites = (- self.correction_matrix * self.codes['points']).groupby(level=0, axis=1).sum()

        # Au cas où certaines questions sont absentes par absence de pénalités (tous on eu tout bon bon)
        penalites = penalites.reindex(columns=self.init.index, fill_value=0.0)

        self.grades = penalites + self.init
        self.grades[self.grades < 0] = 0

    @property
    def grades_total(self):
        """ Calculate the total grade and format"""
        return pd.concat([self.contacts, self.grades.sum(1).to_frame('points')], axis=1)

    @property
    def grades_rel(self):
        """ Express each question in terms of percentage"""
        return self.grades * 100 / self.totals

    @property
    def mean(self):
        return self.grades.sum(1).mean()

    @property
    def median(self):
        return self.grades.sum(1).median()

    @property
    def error_frequency(self):
        """ Calculate the frequency of each error for each question

        Returns
        -------
        freq_err : DataFrame

        """
        question_order = self.totals.index
        freq_err = self.correction_matrix.sum().to_frame('fréquence erreurs (%)') * 100 / self.contacts.shape[0]
        freq_err = freq_err.sort_values(by=['fréquence erreurs (%)'], ascending=False).reindex(question_order, level=0
                                                                                               ).round(1)
        freq_err = freq_err.join(self.codes['définition'])

        # Filter out "errors" with zero penalty
        freq_err = freq_err.loc[self.codes['points'] != 0]

        return freq_err

    def _get_version(self, student_id):
        try:
            version = self.corr.loc[student_id, 'version']
        except KeyError:
            version = None
        return version

    def give_overview(self, q=None, bins=10, filename=None, fail=50, ):

        if q == 'all':
            for i in self.totals.index:
                print("{} : {}".format(i, self.titles[i]))
                if filename:
                    name = filename + '_' + i
                else:
                    name = None

                out = give_overview(self.grades_rel[[i]], i, bins=bins, filename=name, fail=fail)
        else:
            if q is None:
                grades_q = self.grades_total['points']

            else:
                print("{} : {}".format(q, self.titles[q]))
                grades_q = self.grades_rel[[q]]

            out = give_overview(grades_q, q, bins=bins, filename=filename, fail=fail)
        return out

    def archive_grades(self, directory=None):
        if directory is None:
            directory = ''

        stamp = '_' + datetime.now().isoformat()
        filepath = Path(directory, self.exam_name + '.xlsx')
        timestamped_filepath = Path(directory, self.exam_name + stamp + '.xlsx')


        with pd.ExcelWriter(filepath) as writer:
            self.grades_total.to_excel(writer, sheet_name='notes_totales')
            self.grades.to_excel(writer, sheet_name='detail')
            self.codes.loc[self.correction_matrix.columns].to_excel(writer, sheet_name='codes_pondération')
            self.error_frequency.to_excel(writer, sheet_name='fréquence_erreurs')
            self.correction_matrix.to_excel(writer, sheet_name='erreurs')

        shutil.copyfile(filepath, timestamped_filepath)

    # Compilation de la lettre
    def compilation_message(self, student_id, version=None):

        c = self.correction_matrix.loc[student_id]
        if version:
            question_numbers = self.versions[version].to_dict()
            codes = self.codes.rename(index=question_numbers, level=0).sort_index()
            grades = self.grades.rename(columns=question_numbers, level=0).sort_index()
            totals = self.totals.rename(index=question_numbers, level=0).sort_index()
            titles = self.titles.rename(index=question_numbers, level=0).sort_index()
            c = c.rename(index=question_numbers, level=0)
        else:
            codes = self.codes
            grades = self.grades
            totals = self.totals
            titles = self.titles

        # Index des erreurs commises
        ix = c.where(c != 0.).dropna().sort_index().index

        # Données sur les erreurs commises
        frequence_erreurs = c.loc[ix]
        frequence_erreurs.name = 'Fréquence'
        details_erreurs = pd.concat([codes.loc[ix][['définition', 'points']], frequence_erreurs], axis=1)
        details_erreurs.columns = ['Erreur', 'Points par erreur', 'Fréquence']
        error_table = details_erreurs[['Points par erreur', 'Fréquence', 'Erreur']].to_html()

        lettre = self.message['salutation'].format(self.contacts.loc[student_id, 'prénom'],
                                                   self.contacts.loc[student_id, 'nom'],)
        lettre += self.message['foreword'].replace('\n', '<br>')
        lettre += self.message['score_overview'].format(self.mean, self.median, self.grades.loc[student_id].sum())
        for i in totals.index:
            lettre += self.message['score_details'].format(i, titles[i], grades.loc[student_id, i], totals[i])

        lettre += self.message['mistakes_details'].format(error_table)
        lettre += self.message['closing'].replace('\n', '<br>')
        lettre += _txt_end
        return lettre

    def send_results(self, sender, server, targeted_recipients=None, bcc_recipients=None, cc_recipients=None,
                     reply_to=None, exam_dir=None):
        """ Send result report to each student


        """


        do_send = input('Ready to send emails? [y/n]')

        if do_send == 'y':
            server_login = getpass.getpass('Insert your server login ID (e.g., p-matricule): ')
            password = getpass.getpass('Insert your password for the email server: ')


            if bcc_recipients is None:
                bcc_recipients = []

            if cc_recipients is None:
                cc_recipients = []

            if reply_to is None:
                reply_to = []

            if targeted_recipients is None:
                recipients = self.correction_matrix.index
            else:
                recipients = targeted_recipients

            # Contact server
            with smtplib.SMTP_SSL(server) as server:
                server.login(server_login, password)

                # Write email
                for student_id in recipients:

                    version = self._get_version(student_id)
                    msg = MIMEMultipart()

                    msg['Subject'] = "Votre note pour : {}".format(self.exam_name)
                    msg['From'] = sender
                    if reply_to:
                        msg['reply-to'] = ', '.join(reply_to)
                    msg['BCC'] = ', '.join(bcc_recipients)
                    msg['CC'] = ', '.join(cc_recipients)
                    msg['To'] = self.contacts.loc[student_id, 'courriel']

                    msg.attach(MIMEText(self.compilation_message(student_id, version), "html"))

                    if exam_dir:
                        for i in os.listdir(exam_dir / str(student_id)):
                            if re.search('.\.pdf$', i):
                                part = MIMEBase('application', 'octet-stream')
                                with open(exam_dir / str(student_id) / i, 'rb') as file:
                                    part.set_payload(file.read())
                                encoders.encode_base64(part)
                                part.add_header('Content-Disposition',
                                                'attachment; filename="{}"'.format(Path(i).name))
                                msg.attach(part)

                    server.send_message(msg)
                    full_name = self.contacts.loc[student_id, ['prénom', 'nom']]
                    print('Message sent to {} {}'.format(*full_name))
                    time.sleep(5)

                print('Done sending messages!')


def give_overview(grades, question=None, bins=None, filename=None, fail=None):
    """ Donner une vue d'ensemble"""
    if question is not None:
        grades = grades[question]
    grades = grades.dropna()

    median = grades.median()
    mean = grades.mean()
    stdev = grades.std()

    print("moyenne: {:.1f} points".format(mean))
    print("déviation standard: {:.1f} points".format(stdev))
    print("valeur médiane: {:.1f} points".format(median))
    print("Nombre total d'étudiants: {:.0f}".format(grades.shape[0]))
    if fail:
        print("Nombre d'échecs: {:.0f}".format(np.sum(grades < fail)))

    try:
        ax = sns.distplot(grades,
                          bins=bins,
                          rug=True, 
                          kde=False, 
                          rug_kws={"color": 'r'},
                          hist_kws={"alpha": .6},
                          axlabel='Notes', 
                          label=question)
        ax.set(ylabel='Nombre d\'étudiants')
        ax.set_ylim(ymin=0)
        #    ax.set_xlim(xmin=0)

        # Plot verticle line 
        dims = ax.axis()
        plt.vlines(median, dims[2], dims[3], label='médiane', 
                   colors='green')
        plt.vlines(mean, dims[2], dims[3], label='moyenne')
        plt.vlines(mean-stdev, dims[2], dims[3], 
                   label='- déviation standard', colors='gray')
        plt.vlines(mean+stdev, dims[2], dims[3], 
                   label='+ déviation standard', colors='gray')
        if fail:
            plt.vlines(fail, dims[2], dims[3], label='passage', colors='red')

        plt.legend(loc='upper left')

        if filename:
            plt.savefig(filename + '.svg')
            plt.savefig(filename + '.pdf')

        plt.show()
    except ValueError:
        ax = None
        grades=None
        print("La distribution des notes n'a pu être représentée graphiquement.")

    return grades, ax


def cluster_grades(grades, threshold_passing, threshold_a=None, threshold_a_star=None):
    """

    Parameters
    ----------
    grades : Data Series
        Numerical grades to regroup between letter grades

    Returns
    -------
    cotes : DataFrame
        Letter grades, following different paradigms

    """
    # Sort passsing grades
    passing_grades = grades[grades > threshold_passing]
    cotes = grades.sort_values().to_frame('total')

    # Calculate deltas between adjacent grades
    cotes['delta'] = ((cotes['total'] - cotes['total'].shift(+1)).round(1)).fillna(0)

    try:
        # division in quantiles
        cotes['quantiles'] = pd.qcut(passing_grades, q=len(_ALL_PASSING_NORMAL_LETTERS), labels=_ALL_PASSING_NORMAL_LETTERS)

        # division in equal grande ranges
        cotes['equal_ranges'] = pd.cut(passing_grades, bins=len(_ALL_PASSING_LETTERS), labels=_ALL_PASSING_LETTERS)

        # Jenks Natural break clustering
        breaks = jenkspy.jenks_breaks(passing_grades.values, len(_ALL_PASSING_LETTERS))
        cotes['jenks'] = pd.cut(passing_grades, bins=breaks, labels=_ALL_PASSING_LETTERS, include_lowest=True)
    except ValueError:
        print("Cluster analysis failed. Grade distribution too simple, with duplicate edges?")

    # Proposed divisions

    if threshold_a:
        breaks = pd.DataFrame(all_letters(threshold_passing, threshold_a, threshold_a_star))
        breaks = breaks.append([100.0])
        cotes['proposed'] = pd.cut(passing_grades, bins=breaks[0], labels=breaks[1].dropna(), include_lowest=True)

    # Finaliser
    letter_cols = [i for i in cotes.columns if i not in ['mintotal', 'delta']]
    cotes.loc[:, letter_cols] = cotes[letter_cols].astype(object).fillna(_FAIL_LETTER)

    return cotes


def all_letters(threshold_d, threshold_a, threshold_a_star):

    dtoa = ['d ', 'd+', 'c ', 'c+', 'b ', 'b+', 'a ']

    nb_thresholds = len(dtoa) + 1
    step = (threshold_a - threshold_d) / (len(dtoa) - 1)
    thresholds = np.empty((nb_thresholds, 2), dtype='object')
    thresholds = thresholds.reshape((nb_thresholds, 2))
    
    # Calculer les seuils intermédiaires
    i = 0
    for lettre in dtoa:
        thresholds[i, 0] = threshold_d
        thresholds[i, 1] = lettre
        threshold_d += step
        i += 1
    thresholds[-2, 0] = threshold_a  # to avoid rounding errors
    thresholds[-1, :] = [threshold_a_star, 'a*']
    return thresholds

def identify_forcing(grades, threshold_d, threshold_a, threshold_a_star, forcing=0.2):

    thresholds = all_letters(threshold_d, threshold_a, threshold_a_star)

    for i in thresholds:
        diff = i[0] - grades
        bo = (diff > 0.0) & (diff < forcing)
        if bo.any():
            print(f"Potential forcing for {i[1]}:")
            print(grades.loc[bo])
            print(" ")



def give_letter(grade, threshold_d, threshold_a, threshold_a_star, forcing=0.0):

    thresholds = all_letters(threshold_d, threshold_a, threshold_a_star)

    # Calculer la cote pour chaque grade
    letter = 'f'
    for i in range(thresholds.shape[0]):
        if grade >= thresholds[i, 0] - forcing:
            letter = thresholds[i, 1]
        else:
            break
    return letter

# Helper functions


def _clean(s):
    try:
        # white space cleanup
        out = re.sub('\s+', '', s)

        # case specific cleanup
        out = re.sub('\.', '', out)

        # split each error code separately
        out = out.split(',')
    except TypeError:
        out = []
    return out
