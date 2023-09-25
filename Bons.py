#
#    Copyright (c) 2023 Gérard Parat <aero@lunique.fr>
#
""" Création de bons de vol de découverte ou d'initiation."""

import os
import wx
import wx.adv
import wx.html
import configparser
import tempfile
import platform

import wx.lib.agw.pyprogress as PP

from OutilsBons import ID_Generator, genereBon
from threading import Thread
from pathlib import Path
from signal import SIGTERM

class FrameHelp(wx.Frame):
    """Création de la fenêtre d'aide."""

    def __init__(self, parent, title):
        """Initialisation de la fenêtre d'aide."""
        wx.Frame.__init__(self, parent=parent, title=title)
        self.InitTrame()
        self.Show()

    def InitTrame(self):
        """Initialisation de la fenêtre avec le fichier d'aide."""
        sizer = wx.BoxSizer()
        self.SetSizer(sizer)
        helpWin = wx.html.HtmlWindow(self)
        helpWin.LoadFile('help.html')
        sizer.Add(helpWin, 1, wx.EXPAND)
        sizer.Layout()


class FrameConf(wx.Frame):
    """Création de la fenêtre de configuration."""

    def __init__(self, parent, title):
        """Initialisation de la fenêtre de configuration."""
        wx.Frame.__init__(self, parent=parent, title=title)
        self.InitTrame()
        self.Show()

    def InitTrame(self):
        """Initialisation de la fenêtre avec les paramètres."""
        sizer = wx.GridBagSizer(5, 5)
        self.SetSizer(sizer)

        labelModeleVD = wx.StaticText(self, label='Modèle VD')
        sizer.Add(labelModeleVD, pos=(0, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.modeleVD = config['Fichiers']['modelevd']
        self.champModeleVD = wx.TextCtrl(self, value=self.modeleVD)
        sizer.Add(self.champModeleVD, pos=(0,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelClasseurVD = wx.StaticText(self, label='Classeur VD')
        sizer.Add(labelClasseurVD, pos=(1, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.classeurVD = config['Fichiers']['classeurvd']
        self.champClasseurVD = wx.TextCtrl(self, value=self.classeurVD)
        sizer.Add(self.champClasseurVD, pos=(1,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelDossierVD = wx.StaticText(self, label='Dossier VD')
        sizer.Add(labelDossierVD, pos=(2, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.dossierVD = config['Fichiers']['dossiervd']
        self.champDossierVD = wx.TextCtrl(self, value=self.dossierVD)
        sizer.Add(self.champDossierVD, pos=(2,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelModeleVI = wx.StaticText(self, label='Modèle VI')
        sizer.Add(labelModeleVI, pos=(3, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.modeleVI = config['Fichiers']['modelevi']
        self.champModeleVI = wx.TextCtrl(self, value=self.modeleVI)
        sizer.Add(self.champModeleVI, pos=(3,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelClasseurVI = wx.StaticText(self, label='Classeur VI')
        sizer.Add(labelClasseurVI, pos=(4, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.classeurVI = config['Fichiers']['classeurvi']
        self.champClasseurVI = wx.TextCtrl(self, value=self.classeurVI)
        sizer.Add(self.champClasseurVI, pos=(4,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelDossierVI = wx.StaticText(self, label='Dossier VI')
        sizer.Add(labelDossierVI, pos=(5, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.dossierVI = config['Fichiers']['dossiervi']
        self.champDossierVI = wx.TextCtrl(self, value=self.dossierVI)
        sizer.Add(self.champDossierVI, pos=(5,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelCrypto = wx.StaticText(self, label='GnuPG')
        sizer.Add(labelCrypto, pos=(6, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.crypto = config['Crypto']['logiciel']
        self.champCrypto = wx.TextCtrl(self, value=self.crypto)
        sizer.Add(self.champCrypto, pos=(6,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        labelClefID = wx.StaticText(self, label='ID Clef')
        sizer.Add(labelClefID, pos=(7, 0), flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.clef = config['Crypto']['clefid']
        self.champClef = wx.TextCtrl(self, value=self.clef)
        sizer.Add(self.champClef, pos=(7,1), span=(1, 30), flag=wx.LEFT|wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

        self.boutonModeleVD = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonModeleVD, pos=(0, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonClasseurVD = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonClasseurVD, pos=(1, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonDossierVD = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonDossierVD, pos=(2, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonModeleVI = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonModeleVI, pos=(3, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonClasseurVI = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonClasseurVI, pos=(4, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonDossierVI = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonDossierVI, pos=(5, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonCrypto = wx.Button(self, label='Parcourir...')
        sizer.Add(self.boutonCrypto, pos=(6, 31), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT|wx.ALIGN_LEFT, border=5)

        self.boutonAnnuler = wx.Button(self, label='Annuler')
        sizer.Add(self.boutonAnnuler, pos=(9, 0),
            flag=wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT, border=5)

        self.boutonSauver = wx.Button(self, label='Sauver')
        sizer.Add(self.boutonSauver, pos=(9, 31), span=(1, 1),
            flag=wx.BOTTOM|wx.RIGHT|wx.ALIGN_RIGHT, border=5)

        self.boutonModeleVD.Bind(wx.EVT_BUTTON, self.OnBoutonModeleVD)
        self.boutonClasseurVD.Bind(wx.EVT_BUTTON, self.OnBoutonClasseurVD)
        self.boutonDossierVD.Bind(wx.EVT_BUTTON, self.OnBoutonDossierVD)
        self.boutonModeleVI.Bind(wx.EVT_BUTTON, self.OnBoutonModeleVI)
        self.boutonClasseurVI.Bind(wx.EVT_BUTTON, self.OnBoutonClasseurVI)
        self.boutonDossierVI.Bind(wx.EVT_BUTTON, self.OnBoutonDossierVI)
        self.boutonCrypto.Bind(wx.EVT_BUTTON, self.OnBoutonCrypto)
        self.champClef.Bind(wx.EVT_TEXT, self.OnClefID)

        self.boutonAnnuler.Bind(wx.EVT_BUTTON, self.OnBoutonAnnuler)
        self.Bind(wx.EVT_CLOSE, self.OnBoutonAnnuler)

        self.boutonSauver.Bind(wx.EVT_BUTTON, self.OnBoutonSauver)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(8)
        self.Fit()

    def OnBoutonModeleVD(self, event):
        """Recherche du fichier de modèle de vols de découverte."""
        dialogue = wx.FileDialog(self, 'Fichier modèle des bons de vol découverte',
            os.path.dirname(self.modeleVD), wildcard='Modèle de BON|*.docx|Tous les fichiers|*')
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champModeleVD.SetValue(fichier)
            config['Fichiers']['modelevd'] = fichier

    def OnBoutonClasseurVD(self, event):
        """Recherche du fichier de classeur de vols de découverte."""
        dialogue = wx.FileDialog(self, 'Fichier classeur des bons de vol découverte',
            os.path.dirname(self.modeleVD), wildcard='Classeur de BON|*.xlsx|Tous les fichiers|*')
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champClasseurVD.SetValue(fichier)
            config['Fichiers']['classeurvd'] = fichier

    def OnBoutonDossierVD(self, event):
        """Recherche du dossier des vols de découverte."""
        dialogue = wx.DirDialog(self, 'Dossier des bons de vol découverte',
            os.path.dirname(self.dossierVD))
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champDossierVD.SetValue(fichier)
            config['Fichiers']['dossiervd'] = fichier

    def OnBoutonModeleVI(self, event):
        """Recherche du fichier de modèle de vols d'initiation."""
        dialogue = wx.FileDialog(self, 'Fichier modèle des bons de vol d\'initiation',
            os.path.dirname(self.modeleVI), wildcard='Modèle de BON|*.docx|Tous les fichiers|*')
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champModeleVI.SetValue(fichier)
            config['Fichiers']['modelevi'] = fichier

    def OnBoutonClasseurVI(self, event):
        """Recherche du fichier de classeur de vols d'initiation."""
        dialogue = wx.FileDialog(self, 'Fichier classeur des bons de vol d\'initiation',
            os.path.dirname(self.modeleVI), wildcard='Classeur de BON|*.xlsx|Tous les fichiers|*')
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champClasseurVI.SetValue(fichier)
            config['Fichiers']['classeurvi'] = fichier

    def OnBoutonDossierVI(self, event):
        """Recherche du dossier des vols d'initiation."""
        dialogue = wx.DirDialog(self, 'Dossier des bons de vol d\'initiation',
            os.path.dirname(self.dossierVI))
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champDossierVI.SetValue(fichier)
            config['Fichiers']['dossiervi'] = fichier

    def OnBoutonCrypto(self, event):
        """Recherche du logiciel de cryptographie."""
        dialogue = wx.FileDialog(self, 'Logiciel de signature électronique GnuPG',
            os.path.dirname(self.crypto), wildcard='Logiciels|*.exe|Tous les fichiers|*')
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPath()
            self.champCrypto.SetValue(fichier)
            config['Crypto']['logiciel'] = fichier

    def OnClefID(self, event):
        """Configuration de la clef de cryptographie."""
        config['Crypto']['clefid'] = self.champClef.GetLineText(0)

    def OnBoutonAnnuler(self, event):
        """Annulation des éventuelles modifications de la configuration."""
        config.read('config.ini')
        self.Destroy()

    def OnBoutonSauver(self, event):
        """Sauvegarde des éventuelles modifications de la configuration."""
        with open('config.ini', 'w') as configfile:
            config.write(configfile)
        self.Destroy()

class TabVol(wx.Panel):
    """Création de l'onglet."""

    def __init__(self, parent, name, suiteok):
        """Initialisation de l'onglet."""
        wx.Panel.__init__(self, parent)
        self.suiteValide = suiteok
        self.name = name
        self.parent = parent
        self.InitTab()

    def InitTab(self):
        """Initialisation de l'onglet avec les paramètres."""
        logo = config['Images']['logo']
        sizer = wx.GridBagSizer(5, 5)
        self.SetSizer(sizer)

        self.timer = wx.Timer(self)

        self.NumeroBon = ID_Generator(2, 5)
        labelBon = wx.StaticText(self, label='Bon n° ')
        labelNumero = wx.StaticText(self, label=' '+self.NumeroBon+' ')
        labelNumero.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        if self.name == 'vd':
            labelNumero.SetBackgroundColour(wx.Colour(0, 160, 0))
        else:
            labelNumero.SetBackgroundColour(wx.Colour(0, 0, 255))
        labelNumero.SetForegroundColour(wx.Colour(255, 255, 255))
        sizer.Add(labelBon, pos=(0, 0), flag=wx.BOTTOM|wx.ALIGN_BOTTOM|wx.ALIGN_RIGHT, border=10)
        sizer.Add(labelNumero, pos=(0, 1), span=(1, 2), flag=wx.BOTTOM|wx.ALIGN_BOTTOM, border=10)

        logo = wx.StaticBitmap(self, bitmap=wx.Bitmap(logo))
        sizer.Add(logo, pos=(0, 6), span=(1, 1),
            flag=wx.TOP|wx.BOTTOM|wx.RIGHT, border=10)

        ligne1 = wx.StaticLine(self)
        sizer.Add(ligne1, pos=(1, 0), span=(1, 7),
            flag=wx.EXPAND|wx.BOTTOM, border=5)

        labelPax='Passager'
        if self.name == 'vd':
           labelPax += 's'
        textePax = wx.StaticText(self, label=labelPax)
        sizer.Add(textePax, pos=(2, 0), flag=wx.TOP|wx.LEFT, border=5)

        labelNom1 = wx.StaticText(self, label='Nom')
        labelPrenom1 = wx.StaticText(self, label='Prénom')
        if self.name == 'vd':
            labelNom2 = wx.StaticText(self, label='Nom')
            labelPrenom2 = wx.StaticText(self, label='Prénom')
            labelNom3 = wx.StaticText(self, label='Nom')
            labelPrenom3 = wx.StaticText(self, label='Prénom')

        sizer.Add(labelNom1, pos=(3, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.champNom1 = wx.TextCtrl(self)
        sizer.Add(self.champNom1, pos=(3, 2), flag=wx.TOP|wx.EXPAND)
        sizer.Add(labelPrenom1, pos=(3, 3), flag=wx.LEFT|wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.champPrenom1 = wx.TextCtrl(self)
        sizer.Add(self.champPrenom1, pos=(3, 4), span=(1, 1), flag=wx.TOP|wx.EXPAND)

        if self.name == 'vd':
            sizer.Add(labelNom2, pos=(4, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
            self.champNom2 = wx.TextCtrl(self)
            sizer.Add(self.champNom2, pos=(4, 2), flag=wx.TOP|wx.EXPAND)
            sizer.Add(labelPrenom2, pos=(4, 3), flag=wx.LEFT|wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, border=5)
            self.champPrenom2 = wx.TextCtrl(self)
            sizer.Add(self.champPrenom2, pos=(4, 4), flag=wx.TOP|wx.EXPAND)
    
            sizer.Add(labelNom3, pos=(5, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
            self.champNom3 = wx.TextCtrl(self)
            sizer.Add(self.champNom3, pos=(5, 2), flag=wx.TOP|wx.EXPAND)
            sizer.Add(labelPrenom3, pos=(5, 3), flag=wx.LEFT|wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, border=5)
            self.champPrenom3 = wx.TextCtrl(self)
            sizer.Add(self.champPrenom3, pos=(5, 4), flag=wx.TOP|wx.EXPAND)

        boxsizer1 = wx.BoxSizer()
        self.champParent = wx.CheckBox(self, label='Autorisation parentale')
        boxsizer1.Add(self.champParent,
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP|wx.ALIGN_LEFT, border=5)
        sizer.Add(boxsizer1, pos=(6, 0), span=(1, 7),
            flag=wx.EXPAND|wx.TOP|wx.LEFT|wx.RIGHT , border=5)

        ligne2 = wx.StaticLine(self)
        sizer.Add(ligne2, pos=(7, 0), span=(1, 7),
            flag=wx.EXPAND|wx.BOTTOM, border=5)

        textePaiement = wx.StaticText(self, label='Paiement')
        sizer.Add(textePaiement, pos=(8, 0), flag=wx.TOP|wx.LEFT, border=5)
        labelPayeur = wx.StaticText(self, label='Payeur')
        sizer.Add(labelPayeur, pos=(9, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.champPayeur = wx.TextCtrl(self)
        sizer.Add(self.champPayeur, pos=(9, 2), flag=wx.TOP|wx.EXPAND)
        labelDatePaiement = wx.StaticText(self, label='Date')
        sizer.Add(labelDatePaiement, pos=(10, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.choixDatePaiement = wx.adv.DatePickerCtrl(self)
        sizer.Add(self.choixDatePaiement, pos=(10, 2), flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT, border=5)
        self.comboPaiement = wx.ComboBox(self, value='Virement',
            choices=['Virement','Carte bancaire','Espèce','Chèque'], style=wx.CB_READONLY)
        sizer.Add(self.comboPaiement, pos=(11, 1), span=(1, 2),
            flag=wx.EXPAND, border=5)
        self.labelCheque = wx.StaticText(self, label='Numéro')
        sizer.Add(self.labelCheque, pos=(11, 3),
            flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
        self.champCheque = wx.TextCtrl(self)
        sizer.Add(self.champCheque, pos=(11, 4),
            flag=wx.EXPAND|wx.RESERVE_SPACE_EVEN_IF_HIDDEN)
        self.labelBanque = wx.StaticText(self, label='Banque')
        sizer.Add(self.labelBanque, pos=(12, 3),
            flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
        self.champBanque = wx.TextCtrl(self)
        sizer.Add(self.champBanque, pos=(12, 4),
            flag=wx.TOP|wx.EXPAND|wx.RESERVE_SPACE_EVEN_IF_HIDDEN)

        ligne3 = wx.StaticLine(self)
        sizer.Add(ligne3, pos=(13, 0), span=(1, 7),
            flag=wx.EXPAND|wx.BOTTOM, border=5)

        texteClub = wx.StaticText(self, label='Cadre aéro-club')
        sizer.Add(texteClub, pos=(14, 0), flag=wx.TOP|wx.LEFT, border=5)
        if self.name == 'vi':
            self.rboxCours = wx.RadioBox(self, label = 'Vols', choices = ['1', '1+1', '2'],
                majorDimension = 2, style = wx.RA_SPECIFY_COLS,) 
            sizer.Add(self.rboxCours, pos=(15, 0),
                flag=wx.LEFT|wx.TOP|wx.BOTTOM|wx.EXPAND, border=5)
        labelPilote1 = wx.StaticText(self, label='Pilote')
        sizer.Add(labelPilote1, pos=(16, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        listePilotes = []
        if self.name == 'vd':
            liste = list(config.items('Pilotes')) + list(config.items('Instructeurs'))
            liste.remove(config.items('Instructeurs')[0])
            defautPil = config['Pilotes']['pil01']
        else:
            liste = list(config.items('Instructeurs'))
            defautPil = config['Instructeurs']['inst00']
        for item in liste:
            listePilotes.append(item[1])
        self.comboPilote1 = wx.ComboBox(self, value=defautPil,
            choices=listePilotes, style=wx.CB_READONLY)
        sizer.Add(self.comboPilote1, pos=(16, 2),
            flag=wx.EXPAND, border=5)
        labelAvion1 = wx.StaticText(self, label='Avion')
        sizer.Add(labelAvion1, pos=(16, 3), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        listeAvions = []
        liste = list(config.items('Avions'))
        if self.name == 'vd':
            liste = list(config.items('Avions'))
            liste.remove(config.items('Avions')[0])
            defautAvion = config['Avions']['avion01']
        else:
            defautAvion = config['Avions']['avion00']
        for item in liste:
            listeAvions.append(item[1])
        self.comboAvion1 = wx.ComboBox(self, value=defautAvion,
            choices=listeAvions, style=wx.CB_READONLY)
        sizer.Add(self.comboAvion1, pos=(16, 4), span=(1, 2),
            flag=wx.EXPAND, border=5)
        self.labelDate1 = wx.StaticText(self, label='Date')
        sizer.Add(self.labelDate1, pos=(17, 1), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.choixDateVol1 = wx.adv.DatePickerCtrl(self)
        sizer.Add(self.choixDateVol1, pos=(17, 2), flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT, border=5)
        self.labelHeure1 = wx.StaticText(self, label='Heure')
        sizer.Add(self.labelHeure1, pos=(17, 3), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.choixHeureVol1 = wx.adv.TimePickerCtrl(self)
        self.choixHeureVol1.SetTime(15, 0, 0)
        sizer.Add(self.choixHeureVol1, pos=(17, 4), flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT, border=5)
        self.labelTemps1 = wx.StaticText(self, label='Durée')
        sizer.Add(self.labelTemps1, pos=(18, 3), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=5)
        self.choixTempsVol1 = wx.adv.TimePickerCtrl(self)
        self.choixTempsVol1.SetTime(0, 30, 0)
        sizer.Add(self.choixTempsVol1, pos=(18, 4), flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT, border=5)
        if self.name == 'vi':
            self.labelPilote2 = wx.StaticText(self, label='Pilote')
            sizer.Add(self.labelPilote2, pos=(19, 1),
                flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            listePilotes = []
            liste = list(config.items('Instructeurs'))
            defautPil = config['Instructeurs']['inst00']
            for item in liste:
                listePilotes.append(item[1])
            self.comboPilote2 = wx.ComboBox(self, value=defautPil,
                choices=listePilotes, style=wx.CB_READONLY)
            sizer.Add(self.comboPilote2, pos=(19, 2),
                flag=wx.EXPAND|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.labelAvion2 = wx.StaticText(self, label='Avion')
            sizer.Add(self.labelAvion2, pos=(19, 3),
                flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            listeAvions = []
            liste = list(config.items('Avions'))
            defautAvion = config['Avions']['avion00']
            for item in liste:
                listeAvions.append(item[1])
            self.comboAvion2 = wx.ComboBox(self, value=defautAvion,
                choices=listeAvions, style=wx.CB_READONLY)
            sizer.Add(self.comboAvion2, pos=(19, 4), span=(1, 2),
                flag=wx.EXPAND|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.labelDate2 = wx.StaticText(self, label='Date')
            sizer.Add(self.labelDate2, pos=(20, 1),
                flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.choixDateVol2 = wx.adv.DatePickerCtrl(self)
            sizer.Add(self.choixDateVol2, pos=(20, 2),
                flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.labelHeure2 = wx.StaticText(self, label='Heure')
            sizer.Add(self.labelHeure2, pos=(20, 3),
                flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.choixHeureVol2 = wx.adv.TimePickerCtrl(self)
            self.choixHeureVol2.SetTime(15, 0, 0)
            sizer.Add(self.choixHeureVol2, pos=(20, 4),
                flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.labelTemps2 = wx.StaticText(self, label='Durée')
            sizer.Add(self.labelTemps2, pos=(21, 3),
                flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)
            self.choixTempsVol2 = wx.adv.TimePickerCtrl(self)
            self.choixTempsVol2.SetTime(0, 30, 0)
            sizer.Add(self.choixTempsVol2, pos=(21, 4),
                flag=wx.TOP|wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT|wx.RESERVE_SPACE_EVEN_IF_HIDDEN, border=5)

        self.boutonDossier = wx.Button(self, label='Dossier')
        sizer.Add(self.boutonDossier, pos=(23, 0),
            flag=wx.BOTTOM|wx.LEFT|wx.ALIGN_LEFT, border=10)

        self.boutonValider = wx.Button(self, label='Valider')
        sizer.Add(self.boutonValider, pos=(23, 1), span=(1, 5),
            flag=wx.BOTTOM|wx.ALIGN_CENTER, border=10)

        self.boutonEffacer = wx.Button(self, label='Effacer')
        sizer.Add(self.boutonEffacer, pos=(23, 6), span=(1, 1),
            flag=wx.BOTTOM|wx.RIGHT|wx.ALIGN_RIGHT, border=10)

        sizer.AddGrowableCol(2)
        sizer.AddGrowableCol(4)
        sizer.AddGrowableRow(22)

        if self.name == 'vi':
            self.rboxCours.Bind(wx.EVT_RADIOBOX, self.OnChoixCours)
        self.comboPaiement.Bind(wx.EVT_COMBOBOX, self.OnChoixPaiement)
        self.boutonDossier.Bind(wx.EVT_BUTTON, self.OnBoutonDossier)
        self.boutonDossier.Bind(wx.EVT_ENTER_WINDOW, self.OverBoutonDossier)
        self.boutonDossier.Bind(wx.EVT_LEAVE_WINDOW, self.LeaveBoutonDossier)
        self.boutonValider.Bind(wx.EVT_BUTTON, self.OnBoutonValider)
        self.boutonValider.Bind(wx.EVT_ENTER_WINDOW, self.OverBoutonValider)
        self.boutonValider.Bind(wx.EVT_LEAVE_WINDOW, self.LeaveBoutonValider)
        self.boutonEffacer.Bind(wx.EVT_BUTTON, self.OnBoutonEffacer)
        self.boutonEffacer.Bind(wx.EVT_ENTER_WINDOW, self.OverBoutonEffacer)
        self.boutonEffacer.Bind(wx.EVT_LEAVE_WINDOW, self.LeaveBoutonEffacer)

        # Masquage des champs optionnels
        self.labelCheque.Hide()
        self.champCheque.Hide()
        self.labelBanque.Hide()
        self.champBanque.Hide()
        if self.name == 'vi':
            self.labelPilote2.Hide()
            self.comboPilote2.Hide()
            self.labelAvion2.Hide()
            self.comboAvion2.Hide()
            self.labelDate2.Hide()
            self.choixDateVol2.Hide()
            self.labelHeure2.Hide()
            self.choixHeureVol2.Hide()
            self.labelTemps2.Hide()
            self.choixTempsVol2.Hide()
        
        self.Layout()

    def OnChoixPaiement(self, event):
        """Activation selon le choix de mode de paiement."""
        choix = self.comboPaiement.GetStringSelection()
        if choix == 'Chèque':
            self.labelCheque.Show()
            self.champCheque.Show()
            self.labelBanque.Show()
            self.champBanque.Show()
        else:
            self.labelCheque.Hide()
            self.champCheque.Hide()
            self.labelBanque.Hide()
            self.champBanque.Hide()
        self.Layout()

    def OnChoixCours(self, event):
        """Activation selon le choix du nombre de vols."""
        choix = self.rboxCours.GetStringSelection()
        if choix == '2':
            self.labelPilote2.Show()
            self.comboPilote2.Show()
            self.labelAvion2.Show()
            self.comboAvion2.Show()
            self.labelDate2.Show()
            self.choixDateVol2.Show()
            self.labelHeure2.Show()
            self.choixHeureVol2.Show()
            self.labelTemps2.Show()
            self.choixTempsVol2.Show()
        else:
            self.labelPilote2.Hide()
            self.comboPilote2.Hide()
            self.labelAvion2.Hide()
            self.comboAvion2.Hide()
            self.labelDate2.Hide()
            self.choixDateVol2.Hide()
            self.labelHeure2.Hide()
            self.choixHeureVol2.Hide()
            self.labelTemps2.Hide()
            self.choixTempsVol2.Hide()
        self.Layout()

    def OnBoutonDossier(self, event):
        """Visualisation des dossiers de vol."""
        if self.name == 'vd':
            dialogue = wx.FileDialog(self, 'Dossier des bons de vol de découverte',config['Fichiers']['dossiervd'],
                wildcard='Fichiers de BON|*.xlsx;*.docx;*.pdf|Tous les fichiers|*', style=wx.FD_MULTIPLE)
        else:
            dialogue = wx.FileDialog(self, 'Dossier des bons de vol d\'initiation',config['Fichiers']['dossiervi'],
                wildcard='Fichiers de BON|*.xlsx;*.docx;*.pdf|Tous les fichiers|*', style=wx.FD_MULTIPLE)
        if dialogue.ShowModal() == wx.ID_OK:
            fichier = dialogue.GetPaths()
            wx.LaunchDefaultApplication(fichier[0])

    def OverBoutonDossier(self, event):
        """Activation sur la présence de la souris."""
        self.boutonDossier.SetBackgroundColour(wx.Colour(220, 220, 220))
        event.Skip()

    def LeaveBoutonDossier(self, event):
        """Activation sur l'absence de la souris."""
        self.boutonDossier.SetBackgroundColour(wx.NullColour)
        event.Skip()

    def OnBoutonValider(self, event):
        """Lancement de la création du bon de vol."""
        # Génération du bon de vol uniquement avec une suite bureautique valide 
        if self.suiteValide:
            # Génération du bon de vol
            if not self.timer.IsRunning():
                # Barre de progression
                progressBar = PP.PyProgress(None, wx.ID_ANY, 'Bon de vol',
                                    'Génération en cours',
                                    agwStyle=wx.PD_APP_MODAL)
                progressBar.SetGaugeProportion(0.2)
                progressBar.SetGaugeSteps(50)
                progressBar.SetGaugeBackground(wx.WHITE)
                progressBar.SetFirstGradientColour(wx.WHITE)
                progressBar.SetSecondGradientColour(wx.BLUE)
                # Temporisation de 30 s en cas d'échec
                self.timer.StartOnce(30000)
                # Préparation du bon de vol
                bon = self.NumeroBon
                nom1 = self.champNom1.GetLineText(0)
                prenom1 = self.champPrenom1.GetLineText(0)
                typebon = self.name
                if typebon == 'vd':
                    nom2 = self.champNom2.GetLineText(0)
                    prenom2 = self.champPrenom2.GetLineText(0)
                    nom3 = self.champNom3.GetLineText(0)
                    prenom3 = self.champPrenom3.GetLineText(0)
                else:
                    nom2 = ''
                    prenom2 = ''
                    nom3 = ''
                    prenom3 = ''
                autoparent = self.champParent.GetValue()
                payeur = self.champPayeur.GetLineText(0)
                datePaiement = self.choixDatePaiement.GetValue()
                choixPaiement = self.comboPaiement.GetStringSelection()
                if choixPaiement == 'Chèque':
                    numcheque = self.champCheque.GetLineText(0)
                    banque = self.champBanque.GetLineText(0)
                else:
                    numcheque = ''
                    banque = ''
                pilote1 = self.comboPilote1.GetStringSelection()
                avion1 = self.comboAvion1.GetStringSelection()
                date1 = self.choixDateVol1.GetValue()
                heure1 = self.choixHeureVol1.GetValue()
                temps1 = self.choixTempsVol1.GetValue()
                aujourdhui = wx.DateTime.Now().GetDateOnly()
                if typebon == 'vi':
                    pilote2 = self.comboPilote2.GetStringSelection()
                    avion2 = self.comboAvion2.GetStringSelection()
                    date2 = self.choixDateVol2.GetValue()
                    heure2 = self.choixHeureVol2.GetValue()
                    temps2 = self.choixTempsVol2.GetValue()
                    cours = self.rboxCours.GetStringSelection()
                    if cours == '1':
                        tarif = config.getint('Tarifs', 'vol1')
                    elif cours == '1+1':
                        tarif = config.getint('Tarifs', 'vol2')
                    else:
                        tarif = config.getint('Tarifs', 'vol1') + config.getint('Tarifs', 'vol2')
                    if cours == '1' or cours == '1+1':
                        # Mise à zéro si on a une date de vol le même jour...
                        if date1.GetDateOnly() == aujourdhui:
                            pilote1 = ''
                            avion1 = ''
                            date1 = ''
                            heure1 = ''
                            temps1 = ''
                        pilote2 = ''
                        avion2 = ''
                        date2 = ''
                        heure2 = ''
                        temps2 = ''
                    else:
                        # Mise à zéro si on a une date de vol le même jour...
                        if date1.GetDateOnly() == aujourdhui:
                            pilote1 = ''
                            avion1 = ''
                            date1 = ''
                            heure1 = ''
                            temps1 = ''
                        # Mise à zéro si on a une date de vol le même jour...
                        if date2.GetDateOnly() == aujourdhui:
                            pilote2 = ''
                            avion2 = ''
                            date2 = ''
                            heure2 = ''
                            temps2 = ''
                else:
                    tarif = None
                    cours = ''
                    pilote2 = ''
                    avion2 = ''
                    date2 = ''
                    heure2 = ''
                    temps2 = ''
                suiteOffice = config['Suite Office']['logiciel']
                cheminLibreOffice = config['Suite Office']['libreoffice']
                cheminGPG = config['Crypto']['logiciel']
                clef = config['Crypto']['clefid']
                debug = config['Crypto'].getboolean('debug')
                cheminVD = config['Fichiers']['dossiervd']
                classeurVD = config['Fichiers']['classeurvd']
                modeleVD = config['Fichiers']['modelevd']
                cheminVI = config['Fichiers']['dossiervi']
                classeurVI = config['Fichiers']['classeurvi']
                modeleVI = config['Fichiers']['modelevi']
                # Exécution de la génération du bon de vol dans un fil
                finTravail = [False]
                travail = Thread(target=genereBon, args=(
                    bon, nom1, prenom1, typebon, nom2, prenom2, nom3, prenom3,
                    autoparent, choixPaiement, payeur, datePaiement, numcheque, banque,
                    date1, heure1, temps1, pilote1, avion1, cours,
                    date2, heure2, temps2, pilote2, avion2, tarif,
                    cheminVD, classeurVD, modeleVD, cheminVI, classeurVI, modeleVI,
                    suiteOffice, cheminLibreOffice, cheminGPG, clef, debug, finTravail)
                    )
                travail.start()
                # Attente de la fin d'exécution
                while not finTravail[0]:
                    wx.MilliSleep(30)
                    progressBar.UpdatePulse()
                progressBar.Destroy()

    def OverBoutonValider(self, event):
        """Activation sur la présence de la souris."""
        self.boutonValider.SetBackgroundColour(wx.Colour(180, 255, 180))
        event.Skip()

    def LeaveBoutonValider(self, event):
        """Activation sur l'absence de la souris."""
        self.boutonValider.SetBackgroundColour(wx.NullColour)
        event.Skip()

    def OnBoutonEffacer(self, event):
        """Remise à zéro des paramètres de la fenêtre."""
        if self.name == 'vd':
            index = 0
            label = 'Vols de découverte'
        else:
            index = 1
            label = 'Vols d\'initiation'
        self.parent.DeletePage(index)
        self.__init__(self.parent, self.name, self.suiteValide)
        self.parent.InsertPage(index, self, label, True)

    def OverBoutonEffacer(self, event):
        """Activation sur la présence de la souris."""
        self.boutonEffacer.SetBackgroundColour(wx.Colour(255, 200, 200))
        event.Skip()

    def LeaveBoutonEffacer(self, event):
        """Activation sur l'absence de la souris."""
        self.boutonEffacer.SetBackgroundColour(wx.NullColour)
        event.Skip()

class MainFrame(wx.Frame):
    """Fenêtre principale avec les menus et les onglets."""

    def __init__(self):
        """Initialisation de la fenêtre principale."""

        self.ConfigUI()
        
        wx.Frame.__init__(self, None, title=self.titre)

        self.serveurUno()
        self.InitUI()
        self.Centre()
        self.Show()

    def InitUI(self):
        """Construction des menus et les onglets."""

        panel = wx.Panel(self)
        notebook = wx.Notebook(panel)
        menubar = wx.MenuBar()

        fileMenu = wx.Menu()
        quitItem = fileMenu.Append(wx.ID_EXIT, '&Quitter', 'Quitter l\'application')
        confItem = fileMenu.Append(wx.ID_PREFERENCES, '&Préférences', 'Configurer l\'application')
        menubar.Append(fileMenu, '&Fichier')

        helpMenu = wx.Menu()
        helpItem = helpMenu.Append(wx.ID_HELP, '&Aide', 'Aide')
        aboutItem = helpMenu.Append(wx.ID_ABOUT, 'A &propos', 'Informations')
        menubar.Append(helpMenu, '&Aide')

        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_CLOSE, self.onClose)
        self.Bind(wx.EVT_MENU, self.OnQuit, quitItem)
        self.Bind(wx.EVT_MENU, self.OnSettings, confItem)
        self.Bind(wx.EVT_MENU, self.OnHelp, helpItem)
        self.Bind(wx.EVT_MENU, self.OnAboutBox, aboutItem)

        # Création des onglets.
        tabvd = TabVol(notebook, name='vd', suiteok=self.suiteValide)
        tabvi = TabVol(notebook, name='vi', suiteok=self.suiteValide)

        # Ajout des onglets à la fenêtre avec leur nom.
        notebook.AddPage(tabvd, self.titreVD)
        notebook.AddPage(tabvi, self.titreVI)

        sizer = wx.BoxSizer()
        sizer.Add(notebook, 1, wx.EXPAND)
        panel.SetSizer(sizer)
        sizer.Fit(self)

    def onClose(self, event):
        """Fermeture de la fenêtre principale."""
        # On stoppe le serveur Uno
        if self.pid:
            os.kill(self.pid, SIGTERM)
        event.Skip()

    def OnQuit(self, event):
        """Fermeture de la fenêtre principale."""
        # On stoppe le serveur Uno
        if self.pid:
            os.kill(self.pid, SIGTERM)
        self.Destroy()

    def OnSettings(self, event):
        """Initialisation de la fenêtre de configuration."""
        FrameConf(self, self.trameConf)

    def OnHelp(self, event):
        """Initialisation de la fenêtre de d'aide."""
        FrameHelp(self, self.trameAide)

    def OnAboutBox(self, event):

        description = """Ce logiciel permet la création de bons authentiques
pour les vols de découverte ou d'initiation."""

        licence = """\
GNU AFFERO GENERAL PUBLIC LICENSE
Version 3, 19 November 2007

Copyright (C) 2007 Free Software Foundation, Inc. <https://fsf.org/>
Everyone is permitted to copy and distribute verbatim copies
of this license document, but changing it is not allowed.
<https://www.gnu.org/licenses/>."""
        
        info = wx.adv.AboutDialogInfo()

        logo = config['Images']['logo']
        info.SetIcon(wx.Icon(logo))
        info.SetName('Bons ACD')
        info.SetVersion('2.2')
        info.SetDescription(description)
        info.SetCopyright('(C) 2023 Gérard Parat')
        info.SetWebSite('http://www.aero-club-dreux.com')
        info.SetLicence(licence)
        info.AddDeveloper('Gérard Parat')

        wx.adv.AboutBox(info)

    def ConfigUI(self):
        """Lecture des paramètres de configuration."""
        global config
        config = configparser.ConfigParser()
        config.read('config.ini')
        self.titre = config['Titres']['principal']
        self.titreVD = config['Titres']['tabvd']
        self.titreVI = config['Titres']['tabvi']
        self.trameAide = config['Titres']['trameaide']
        self.trameConf = config['Titres']['trameconf']

    def serveurUno(self):
        """Lancement du serveur Uno pour LibreOffice"""
        suiteOffice = config['Suite Office']['logiciel']
        self.suiteValide = False
        self.pid = None
        # Importation des modules pour LibreOffice
        if suiteOffice == 'LibreOffice':
            from unoserver.server import UnoServer
            # Lancement du serveur sous Windows
            if platform.system() == 'Windows':
                with tempfile.TemporaryDirectory() as tmpuserdir:
                    tmp_dir = Path(tmpuserdir).as_uri()
                # On lance en dehors du Context Manager
                serveur = UnoServer(user_installation=tmp_dir)
                process = serveur.start(executable='soffice.exe')
                self.pid = process.pid
                self.suiteValide = True
            # Lancement du serveur sous Linux
            elif platform.system() == 'Linux':
                with tempfile.TemporaryDirectory() as tmpuserdir:
                    tmp_dir = Path(tmpuserdir).as_uri()
                    # On lance dans le Context Manager
                    serveur = UnoServer(user_installation=tmp_dir)
                    process = serveur.start()
                    self.pid = process.pid
                    self.suiteValide = True
        # Rien à faire pour Microsoft Office
        elif suiteOffice == 'Microsoft Office':
            if platform.system() == 'Windows':
                self.suiteValide = True
            elif platform.system() == 'Linux':
                # Pas de Microsoft Office sous Linux
                print('La suite bureautique %s n\'est pas supportée !' % suiteOffice)
        # Erreur de suite bureautique...
        else:
            print('La suite bureautique %s n\'est pas supportée !' % suiteOffice)

def main():

    app = wx.App()
    MainFrame()
    app.MainLoop()

if __name__ == '__main__':
    main()
