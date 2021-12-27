import os
import logging
from datetime import datetime
import time
import re
import pandas as pd
import numpy as np
import shutil
from pathlib import Path
import seaborn as sns
import matplotlib.pyplot as plt

import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders


# Set / fix styles
sns.set_style()
pd.DataFrame._repr_latex_ = lambda self: """\centering{}""".format(self.to_latex())


# More formatting / style
#def _repr_latex_(self):
#    return "\centering{}".format(self.to_latex())
#pd.DataFrame._repr_latex_ = _repr_latex_  # monkey patch pandas DataFrame

def correction_parser(filename, exam_name):
    """ A parser to define a `Grader` object based on a spreadsheet template of corrections, correction codes, etc."""
    with pd.ExcelFile(filename) as f:
        raw = pd.read_excel(f, sheet_name='Corrections', header=0, index_col=0)
        raw_codes = pd.read_excel(f, sheet_name='codes', header=0, index_col=[0, 1])
        codes_universels = pd.read_excel(f, sheet_name='codes_universels', header=0, index_col=0)
        totaux = pd.read_excel(f, sheet_name='totaux', header=0, index_col=0).squeeze()
        init = pd.read_excel(f, sheet_name='init', header=0, index_col=0).squeeze()
        versions = pd.read_excel(f, sheet_name='versions', header=0, index_col=0).squeeze()

    return Grader(exam_name, raw, raw_codes, codes_universels, totaux, init, versions)


class Grader:

    def __init__(self, exam_name, raw_corr, raw_codes, codes_universels, totaux, init, versions=None):

        # Data from tremplate
        self.exam_name = exam_name
        self.raw_corr = raw_corr
        self.raw_codes = raw_codes
        self.codes_universels = codes_universels
        self.totaux = totaux
        self.init = init
        self.versions = versions

        # Constants - hardcoded
        self._OK_TERMS = ['ok', '0', '', 'OK']
        self._CONTACT_COLS = ['prénom', 'nom', 'courriel']
        self._COLS_TO_ALWAYS_DROP = ['version', 'Exemple/explication']


        # Refined data
        self.contacts = raw_corr[self._CONTACT_COLS]
        self.corr = raw_corr.drop(self._CONTACT_COLS, axis=1)

        # Semi-final variables
        self.C = None
        self.codes = None

        # Variables finales
        self.notes = None

        self.message = dict()

        self.message['salutation'] = """ Bonjour {} {},

        """

        self.message['foreword'] = ""


        self.message['score overview'] = """

        La moyenne du groupe est de {:.1f}%, et sa note médiane est de {:.1f}%. Vous avez obtenu une note de {:.1f}%. 

        ---------------------

        Voici le détail de vos points: 

        """

        self.message['score_details'] = """ 
            {} : {} points sur {}
            """

        self.message['mistakes_details'] = """

        ---------------------

        Et voici le détail des points perdus:

        {}

        ------------------ """

        self.message['closing'] = """ """

    def calc_grades(self, cols_to_drop=None):
        if cols_to_drop is None:
            cols_to_drop = []

        self._pivot_corr(cols_to_drop=cols_to_drop)
        self._clean_codes(cols_to_drop=cols_to_drop)
        self._check_sanity_and_harmonize()
        self._calc_notes()

    def _pivot_corr(self, cols_to_drop=None):

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

        C = pd.DataFrame(0.0, index=self.corr.index, columns=mcx)

        # Fill
        for ix in self.corr.index:
            for cx in self.corr.columns:
                erreurs = _clean(self.corr.loc[ix, cx])
                for err in erreurs:
                    C.loc(axis=0)[ix].loc[[cx], err] += 1

        # Remove other stuff
        if cols_to_drop is not None:
            C = C.drop(cols_to_drop, axis=1, level=0, errors='ignore')

        C = C.reindex(columns=self.totaux.index, level=0)

        # Enlève OK
        self.C = C.drop(labels=self._OK_TERMS, axis=1, level=1, errors='ignore')

    def _clean_codes(self, cols_to_drop):

        # Drop extraneous columns
        cols_to_drop = self._COLS_TO_ALWAYS_DROP + cols_to_drop
        codes = self.raw_codes.drop(cols_to_drop, axis=1, errors='ignore')

        # Drop empty rows
        todrop = codes['points'].isna()
        self.codes = codes.loc[~todrop]

        # Scale and insert universal codes
        for ix, v in self.totaux.items():
            scaled_codes = (self.codes_universels[['pénalités relatives', 'pénalités_absolues']] * [v, 1])
            selected_codes = scaled_codes.dropna(axis=0, how='all').min(axis=1, skipna=True).to_frame('points')
            new_codes = pd.concat([self.codes_universels['définition'], selected_codes], axis=1, join='inner')
            new_codes.index = pd.MultiIndex.from_product([[ix], new_codes.index])
            self.codes = pd.concat([self.codes, new_codes])
        self.codes.reindex(index=self.totaux.index, level=0)

    def _distribute_codes(self, common_ix, applicables):
        common_codes = self.codes.loc[common_ix]
        tmp = pd.concat([common_codes, ] * len(applicables), keys=applicables)
        codes = pd.concat([self.codes, tmp], axis=0).sort_index()
        self.codes = codes.drop(common_ix, axis=0)

    def _check_sanity_and_harmonize(self):
        missing_codes = self.C.columns.difference(self.codes.index)
        if len(missing_codes):
            logging.warning("Certains codes de correction ne sont pas définis: {}".format(missing_codes))
        else:
            self.codes = self.codes.reindex(self.C.columns)

    def _calc_notes(self):
        penalites = (- self.C * self.codes['points']).groupby(level=0, axis=1).sum()

        # Au cas où certaines questions sont absentes par absence de pénalités (tous on eu tout bon bon)
        penalites = penalites.reindex(columns=self.init.index, fill_value=0.0)

        self.notes = penalites + self.init
        self.notes[self.notes < 0] = 0
        self.mean = self.notes.sum(1).mean()
        self.median = self.notes.sum(1).median()

    @property
    def notes_total(self):
        return pd.concat([self.contacts, self.notes.sum(1).to_frame('points')], axis=1)

    def _get_version(self, pmat):
        try:
            version = self.corr.loc[pmat, 'version']
        except KeyError:
            version = None
        return version

    def give_overview(self, q=None, bins=10, filename=None, fail=50, ):

        # TODO: fix properly
        if q is None:
            notes_q = self.notes_total['points']
        else:
            notes_q = (self.notes / self.totaux)[[q]] * 100

        out = give_overview(notes_q, q, bins=bins, filename=filename, fail=fail)
        return out

    def archive_grades(self, directory=None):
        if directory is None:
            directory = ''

        stamp = '_' + datetime.now().isoformat()
        filepath = Path(directory, self.exam_name + '.xlsx')
        timestamped_filepath = Path(directory, self.exam_name + stamp + '.xlsx')

        with pd.ExcelWriter(filepath) as writer:
            self.notes_total.to_excel(writer, sheet_name='notes_totales')
            self.notes.to_excel(writer, sheet_name='detail')
            self.codes.loc[self.C.columns].to_excel(writer, sheet_name='codes_pondération')
            self.C.to_excel(writer, sheet_name='erreurs')

        shutil.copyfile(filepath, timestamped_filepath)

    # Compilation de la lettre
    def compilation_message(self, pmat, version=None):

        c = self.C.loc[pmat]
        if version:
            question_numbers = self.versions[version].to_dict()
            codes = self.codes.rename(index=question_numbers, level=0).sort_index()
            notes = self.notes.rename(columns=question_numbers, level=0).sort_index()
            totaux = self.totaux.rename(index=question_numbers, level=0).sort_index()
            c = c.rename(index=question_numbers, level=0)
        else:
            codes = self.codes
            notes = self.notes
            totaux = self.totaux

        # Index des erreurs commises
        ix = c.where(c != 0.).dropna().sort_index().index

        # Données sur les erreurs commises
        details_erreurs = codes.loc[ix][['définition', 'points']]
        details_erreurs.columns = ['Erreur', 'points']
        error_table = details_erreurs[['points', 'Erreur']].reset_index().to_markdown(showindex=False,
                                                                                      tablefmt='presto')

        lettre = self.message['salutation'].format(self.contacts.loc[pmat, 'prénom'],
                                                   self.contacts.loc[pmat, 'nom'],)
        lettre += self.message['foreword']
        lettre += self.message['score_overview'].format( self.mean, self.median, self.notes.loc[pmat].sum())
        for i in totaux.index:
            lettre += self.message['score_details'].format(i, notes.loc[pmat, i], totaux[i])

        lettre += self.message['mistakes_details'].format(error_table, details_erreurs['points'])
        lettre += self.message['closing']
        return lettre

    def send_results(self, sender, server, server_login, targeted_recipients=None, bcc_recipients=None, exam_dir=None):

        do_send = input('Ready to send emails? [y/n]')

        if do_send == 'y' and False:
            password = input('Insert your password for the email server.')

            # Contact server
            server = smtplib.SMTP(server, 587)
            server.starttls()
            server.login(server_login, password)

            if targeted_recipients is None:
                recipients = self.C.index
            else:
                recipients = targeted_recipients

            # Write email
            for pmat in recipients:

                version = self._get_version(pmat)
                msg = MIMEMultipart()

                msg['Subject'] = "Votre note pour : {}".format(self.exam_name)
                msg['From'] = sender
                msg['BCC'] = ', '.join(bcc_recipients)
                msg['To'] = self.contacts.loc[pmat, 'courriel']

                msg.attach(MIMEText(self.compilation_message(pmat, version)))

                if exam_dir:
                    for i in os.listdir(exam_dir / str(pmat)):
                        if re.search('.\.pdf$', i):
                            part = MIMEBase('application', 'octet-stream')
                            with open(exam_dir / str(pmat) / i, 'rb') as file:
                                part.set_payload(file.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition',
                                            'attachment; filename="{}"'.format(Path(i).name))
                            msg.attach(part)

                server.send_message(msg)
                print('Message sent to {}'.format(msg['To']))
                time.sleep(5)

            print('Done sending messages!')
            server.quit()



def give_overview(notes, question=None, bins=None, filename=None, fail=None):
    """ Donner une vue d'ensemble"""
    if question is not None:
        notes = notes[question]
    notes = notes.dropna()

    median = notes.median()
    mean = notes.mean()
    stdev = notes.std()

    print("moyenne: {:.1f}%".format(mean))
    print("déviation standard: {:.1f} points de pourcentage".format(stdev))
    print("valeur médiane: {:.1f}%".format(median))
    print("Nombre total d'étudiants: {:.0f}".format(notes.shape[0]))
    if fail:
        print("Nombre d'échecs: {:.0f}".format(np.sum(notes < fail)))
    ax = sns.distplot(notes,
                      bins=bins,
                      rug=True, 
                      kde=False, 
                      rug_kws = {"color":'r'}, 
                      hist_kws= {"alpha":.6}, 
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
        plt.vlines(fail, dims[2], dims[3], label='passage', 
               colors='red')

    plt.legend(loc='upper left')

    if filename:
        plt.savefig(filename + '.svg')
        plt.savefig(filename + '.pdf')

    return notes, ax

def all_letters(seuilD, seuilA, seuilAA):

    dtoa = ['d ', 'd+', 'c ', 'c+', 'b ', 'b+', 'a ']
    step = (seuilA - seuilD) / (len(dtoa) - 1)
    seuils = np.empty((8,2), dtype='object')
    seuils = seuils.reshape((8,2))
    
    # Calculer les seuils intermédiaires
    i = 0
    for lettre in dtoa:
        seuils[i, 0] = seuilD
        seuils[i, 1] = lettre
        seuilD += step
        i += 1
    seuils[-2, 0] = seuilA # to avoid rounding errors
    seuils[-1, :] = [seuilAA, 'a*']
    return seuils

def give_letter(grade, seuilD, seuilA, seuilAA, forcage=0.0, verbose=False):

    seuils = all_letters(seuilD, seuilA, seuilAA)

    # Calculer la cote pour chaque grade
    letter = 'f'
    for i in range(seuils.shape[0]):
        if grade >= seuils[i, 0] - forcage:
            letter = seuils[i, 1]
        else:
            break
    return letter

        



## Helper functions

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
