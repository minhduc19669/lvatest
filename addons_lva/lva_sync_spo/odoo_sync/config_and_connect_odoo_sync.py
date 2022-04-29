"""
Diese Datei ist Bestandteil der Dateisyncronisation
zwischen odoo und SPO. Dabei geht es darum, zu einem
entsprechenden Objekt im Odoo ein entsprechendes Docsets
anzulegen und vorhandene dateien dort abzulegen.

Sie stellt die Funktionen bereit, um
die Konfiguration für den Sync der Maschinenliste,
Angebote und Verkaufsprojekte
zu ermitteln und die entsprechenden Connection-Objekte für
die Zugriffe zurückgeben.

Das Betrifft den Zugriff auf die Odoo-Datenbank, den
Odoo-Filestore und auf die Shapointbibliotheken für die
Ablage der Anhänge.
"""

import sys
import configparser
from odoo.addons.lva_sync_spo.odoo_sync.odoo_connect import *
from odoo.addons.lva_sync_spo.odoo_sync.sftp_filemove import *
from odoo.addons.lva_sync_spo.odoo_sync.spo_connect import *
import tempfile
import os


class getConfAndObjects:
    """
    Klasse zum Einlesen der Configdatei und zur Rückgabe
    von Konfigurationsdaten und Zugriffsobjekten
    """

    def __init__(
            self, conf_file, spo_machine_conn=True,
            spo_crm_conn=True, spo_sales_statistic=False):
        """
        Stellt die Verbindung zur Configdatei her
        
        conf_file               = Pfadangabe für das Config-File
        spo_machine_conn (opt.) = Giebt an, ob SPO-Maschinen-ConnectionObjekte 
                                  erstellt werden sollen
        spo_crm_conn (opt.)     = Giebt an, ob SPO-CRM-ConnectionObjekte 
                                  erstellt werden sollen
        spo_sales_statistic(opt)= Giebt an, ob SPO-Sales-Statistik-
                                  ConnectionsObjekte erstellt werden sollen
        """
        self.__sftp_conn = None
        self.__companies = {}
        self.__spo_create_mach_conn = spo_machine_conn
        self.__spo_create_crm_conn = spo_crm_conn
        self.__config = configparser.ConfigParser()
        self.__config.read(conf_file)
        self.__init_db_connection()
        self.__init_sftp_connection()
        if self.__db_conn_status == True:
            self.__read_companies_list()
            if spo_crm_conn == True:
                self.__init_spo_crm_connection()
            if spo_machine_conn == True:
                self.__init_spo_machine_connection()
        else:
            sys.exit(f'PROGRAM ABGEBROCHEN!\nFehler! Kann keine Vernindung' +
                     f'zur Odoo-Datenbank aufbauen!\n{self.__db_conn_message}')

    def __init_db_connection(self):
        """
        Initialisiert ein Zugriffsobjekt für die odoo-DB.
        für den Zugriff werden die Einstellungen aus 
        der Conf-Datei benutzt.
        
        Es werden die folgenden Variablen und Objekte gesetzt:
        
        self.__db_conn_status   True/False giebt an, ob ein
                                Verbindungs-Objekt zur DB erfolgreich 
                                eingerichtet wurde.
        self.__db_conn_message  Text. Giebt die Fehler-/OK-Meldung für 
                                das DB-Verbindungs-Objekt
        self.__db_conn          Das eigentliche DB-Verbindungs-Objekt
        """

        self.__db_conn = None

        if 'Database' not in self.__config:
            # Wenn kein 'Database'-Abschnitt in der Conf
            self.__db_conn_status = False
            self.__db_conn_message = (
                "Fehler! Der Abschnitt 'Database' fehlt in der Configdatei!")
            return
        else:
            db_conf = self.__config['Database']
            # Hostangabe des Datenbankservers aus Conf. ermitteln
            if 'host' in db_conf:
                db_host = db_conf['host']
            else:
                self.__db_conn_status = False
                self.__db_conn_message = (
                        "Fehler! Die Angabe des Datenbankservers fehlt in " +
                        "der Konfigurationsdatei!")
                return
            # Portangabe (opt.) des datenbankservers aus Conf. ermitteln 
            # oder auf default
            if 'port' in db_conf:
                db_port = int(db_conf['port'])
            else:
                db_port = 5432
            # Benutzer für Datenbank aus Conf. ermitteln
            if 'user' in db_conf:
                db_user = db_conf['user']
            else:
                self.__db_conn_status = False
                self.__db_conn_message = (
                        "Fehler! Die Angabe des Datenbankbenutzers " +
                        "fehlt in der Konfigurationsdatei!")
                return
            # Password für Datenbank aus Conf. ermitteln
            if 'user' in db_conf:
                db_pass = db_conf['pass']
            else:
                self.__db_conn_status = False
                self.__db_conn_message = (
                        "Fehler! Die Angabe des Passwortes fehlt " +
                        "in der Konfigurationsdatei!")
                return
            # Namen für Datenbank aus Conf. ermitteln
            if 'dbname' in db_conf:
                db_name = db_conf['dbname']
            else:
                self.__db_conn_status = False
                self.__db_conn_message = (
                        "Fehler! Die Angabe des Datenbanknamen fehlt " +
                        "in der Konfigurationsdatei!")
                return
            # Optionale Angaben des SSH-Tunnels einlesen
            db_use_tunnel = False
            if 'use_tunnel' in db_conf:
                if db_conf['use_tunnel'] == 'true':
                    db_use_tunnel = True
                    # Tunneldaten einlesen
                    # SSH-Host
                    if 'ssh_host' in db_conf:
                        db_ssh_host = db_conf['ssh_host']
                    else:
                        self.__db_conn_status = False
                        self.__db_conn_message = (
                                "Fehler! In der Konfigdatei ist die Nutzung " +
                                "eines SSH-Tunnels für den Datenbanzugriff " +
                                "angegeben.\n Es fehlt die Angabe des SSH-Hostes!")
                        return
                    # SSH-Port (opt.) ansonsten default
                    if 'ssh_port' in db_conf:
                        db_ssh_port = int(db_conf['ssh_port'])
                    else:
                        db_ssh_port = 22
                    # SSH-Benutzer
                    if 'ssh_user' in db_conf:
                        db_ssh_user = db_conf['ssh_user']
                    else:
                        self.__db_conn_status = False
                        self.__db_conn_message = (
                                "Fehler! In der Konfigdatei ist die Nutzung " +
                                "eines SSH-Tunnels für den Datenbanzugriff" +
                                "angegeben.\n Es fehlt die Angabe des " +
                                "SSH-Benutzers!")
                        return
                    # SSH-Passwort
                    if 'ssh_pass' in db_conf:
                        db_ssh_pass = db_conf['ssh_pass']
                    else:
                        self.__db_conn_status = False
                        self.__db_conn_message = (
                                "Fehler! In der Konfigdatei ist die Nutzung " +
                                "eines SSH-Tunnels für den Datenbanzugriff " +
                                "angegeben.\n Es fehlt die Angabe des " +
                                "SSH-Passwortes!")
                        return

        print("Konfiguration für den Zufgriff auf die Odoo-DB eingelesen.")

        # Objekt für den Datenbankzugriff erstellen
        # !! Hier wird try benutzt, damit bei Fehlermeldungen keine 
        # Accountangaben auf der Ausgabe erscheinen !!
        try:
            self.__db_conn = odoo_connect.OdooDB(
                db_host, db_user, db_pass, db_name, db_port)
            # print(self.__db_conn)
        except:
            self.__db_conn_status = False
            self.__db_conn_message = (
                "Fehler! Fehler bei der Initialisierung des Datenbankobjektes!")
            return

        # Tunneldaten, wenn benötigt setzen
        if db_use_tunnel == True:
            # !! Hier wird try benutzt, damit bei Fehlermeldungen keine
            # Accountangaben auf der Ausgabe erscheinen !!
            # try:
            result = self.__db_conn.ssh_tunnel_init(
                db_ssh_host, db_ssh_user, db_ssh_pass, db_ssh_port)
            # if result[0] == False:
            #    sys.exit(
            #        f'Fehler bei der Initialisierung des ' +
            #        f'SSH-Tunnels des Datenbankobjektes.\n{result[1]}')
            # except:
            #    sys.exit(f'Fehler bei der Initialisierung des ' +
            #             f'SSH-Tunnels des Datenbankobjektes!')

        # Datenbankverbindung testen
        result = self.__db_conn.test_connect()
        if result[0] == False:
            self.__db_conn_status = False
            self.__db_conn_message = (f"Fehler! Fehler bei der Verbindung " +
                                      f"zum Datenbankserver!\n{result[1]}")
            return
        else:
            self.__db_conn_status = True
            self.__db_conn_message = (
                "Die Verbindung zum DB-Server wurde aufgebaut und getestet.")
            return

    def get_db_connection(self):
        """
        Giebt das DB-Zugriffs-Objekt sowie den Status zurück.
        
        Rückgabe: Liste: erster Wert True/False, zweiter Wert Meldung, 
        dritter Wert DB-Zugriffs-Objekt 
        """

        return self.__db_conn_status, self.__db_conn_message, self.__db_conn

    def __init_sftp_connection(self):
        """
        Initialisiert ein Zugriffsobjekt für den Odoo-Filstore.
        für den Zugriff werden die Einstellungen aus 
        der Conf-Datei benutzt.
        
        Es werden die folgenden Variablen und Objekte gesetzt:
        
        self.__sftp_conn_status    True/False giebt an, ob ein
                                Verbindungs-Objekt zur DB erfolgreich 
                                eingerichtet wurde.
        self.__sftp_conn_message  Text. Giebt die Fehler-/OK-Meldung für 
                                das DB-Verbindungs-Objekt
        self.__sftp_conn          Das eigentliche DB-Verbindungs-Objekt
        """

        if 'Filestore' not in self.__config:
            self.__sftp_conn_status = False
            self.__sftp_conn_message = ("Fehler! Kein der Abschnitt " +
                                        "'Filestore' fehlt in der Configdatei!")
            return
        else:
            fs_config = self.__config['Filestore']
            # Zugriffsdaten abfragen
            if 'host' in fs_config:
                fs_host = fs_config['host']
            else:
                self.__sftp_conn_status = False
                self.__sftp_conn_message = (
                    "Fehler! Keine Hostangabe für den Zugriff "
                    "auf den Filestore in der Config!")
                return
            if 'port' in fs_config:
                fs_port = fs_config['port']
            else:
                fs_port = 22
            if 'user' in fs_config:
                fs_user = fs_config['user']
            else:
                self.__sftp_conn_status = False
                self.__sftp_conn_message = (
                        "Fehler! Keine Userangabe für " +
                        "den Zugriff auf den Filestore in der Config!")
            if 'pass' in fs_config:
                fs_pass = fs_config['pass']
            else:
                self.__sftp_conn_status = False
                self.__sftp_conn_message = (
                        "Fehler! Kein Password für den " +
                        "Zufriff auf den Filestore in der Config!")
            if 'base_path' in fs_config:
                fs_base_path = fs_config['base_path']
            else:
                self.__sftp_conn_status = False
                self.__sftp_conn_message = ("Fehler! Kein Basisverzeichnis " +
                                            "für den Filestore in der Config!")

        print("Konfiguration für den Zufgriff auf den Filestore eingelesen.")

        # Objekt für den SFTP-Zugriff
        self.__sftp_conn = sftp_filemove.SSHsftp(
            fs_host, fs_user, fs_pass, fs_port)
        result = self.__sftp_conn.test_conn()
        if result[0] == False:
            self.__sftp_conn_status = False
            self.__sftp_conn_message = (
                    f'Fehler! Kann mit den Configdaten keine Verbindung zum ' +
                    f'SFTP-Server aufnehmen!\n{result[1]}')
            return

        # Basispfade für des SFTP-Objekt setzen
        # self.__tempfolder = tempfile.TemporaryDirectory().name
        # os.mkdir(self.__tempfolder)
        self.__tempfolder = '/home/minhduc/OdooBase/odoo14.local/share/Odoo/filestore/LVA'
        self.__sftp_conn.set_base_path(
            local_path=self.__tempfolder, remote_path=fs_base_path)

        self.__sftp_conn_status = True
        self.__sftp_conn_message = (
            "Die Verbindung zum SFTP-Server wurde aufgebaut und getestet.")
        return

    def get_sftp_connection(self):
        """
        Giebt das SFTP-Zugriffs-Objekt sowie den Status zurück.
        
        Rückgabe: Liste: erster Wert True/False, zweiter Wert Meldung, 
                  dritter Wert DB-Zugriffs-Objekt 
        """

        return (self.__sftp_conn_status,
                self.__sftp_conn_message, self.__sftp_conn)

    def get_sftp_tempfolder(self):
        """
        Giebt des Tempverzeichnis zurück
        """
        return self.__tempfolder

    def __read_companies_list(self):
        """
        Liest die Firmenliste aus der Odoo-DB und generiert
        eine Liste mit ID, Name und Abkürzung
        
        Die Firmenliste wird als self.__companies ausgegeben
        """

        result = self.__db_conn.get_company_list()
        if result[0] == True:
            self.__companies = {}
            for company in result[1]:
                self.__companies[company['id']] = {}
                self.__companies[company['id']]['name'] = company['name']
                if 'short_name' in company:
                    self.__companies[company['id']]['shortname'] = (
                        company['short_name'])
                else:
                    self.__companies[company['id']]['shortname'] = ""

    def get_company_list(self):
        """
        Giebt die Liste der Firmen im Odoo zurück
        """
        return self.__companies

    def __read_spo_machine_config(self):
        """
        Liest die Konfigurationen für die SPO-Verbindungen für 
        die Bibliotheken der Maschinenliste.
        
        Die SPO-Konfigurationen werden als self.__spo_machine_conf[<ID>] 
        ausgegeben. Dabei entspricht die ID der Firmenid im Oddo. 
        Die ID=0 giebt die Defaultwerte an
        """
        # Defaul-Werte für SPO-Konfugurationen ermitteln
        conf_error = False
        if 'SPO' in self.__config:
            spo_default_conf = self.__config['SPO']
            # feststellen, ob Maschinensyncronisation als Default gesetzt ist
            default_masch_sync = True
            if 'sync_machines' in spo_default_conf:
                if spo_default_conf['sync_machines'] == False:
                    default_masch_sync = False

            # Weitere Default-Synceinstell nur ermitteln, wenn Masch-Sync 
            # auf True steht
            if default_masch_sync == True:
                if 'user' in spo_default_conf:
                    default_user = spo_default_conf['user']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe eines Benutzers im Abschnit [SPO]!")
                if 'pass' in spo_default_conf:
                    default_pass = spo_default_conf['pass']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe eines Passworts im Abschnit [SPO]!")
                if 'machine_website' in spo_default_conf:
                    default_website = spo_default_conf['machine_website']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe einer Website im Abschnit [SPO]!")
                if 'machine_library' in spo_default_conf:
                    default_library = spo_default_conf['machine_library']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe einer Biblithek im Abschnit [SPO]!")
                # Feststellen, ob Docset oder Folder
                default_onlyfolder = False
                if 'machine_docset' in spo_default_conf:
                    default_docset = spo_default_conf['machine_docset']
                else:
                    default_onlyfolder = True
                if 'machine_docset_id' in spo_default_conf:
                    default_docset_id = spo_default_conf['machine_docset_id']
                else:
                    default_onlyfolder = True

        # Konfigurationen für die einzelnen Firmen ermitteln
        self.__spo_machine_conf = {}
        for firma in self.__companies:
            self.__spo_machine_conf[firma] = {}
            # Feststellen, ob gesonderte Konfiguration für die entsprechende FA vorhanden
            co_sec = f'SPO_{firma}'
            if co_sec in self.__config:
                # Wenn gesonderte Konfiguration vorhanden
                spo_fa_conf = self.__config[co_sec]
                self.__spo_machine_conf[firma]['sync'] = True
                self.__spo_machine_conf[firma]['message'] = (
                        f'Konfiguration für Firma ' +
                        f'{self.__companies[firma]["name"]} ermittelt')
                # Feststellen, ob Maschinenanhänge nict syncronisiert werden sollen
                if 'sync_machines' in spo_fa_conf:
                    if spo_fa_conf['sync_machines'] == False:
                        self.__spo_machine_conf[firma]['sync'] = False
                        self.__spo_machine_conf[firma]['message'] = (
                                f'Laut Konfiguration soll für Firma ' +
                                f'{self.__companies[firma]["name"]} keine ' +
                                f'Maschinenanhangssynchronisation durchgeführt ' +
                                f'werden.')
                if 'user' in spo_fa_conf:
                    self.__spo_machine_conf[firma]['user'] = spo_fa_conf['user']
                else:
                    try:
                        self.__spo_machine_conf[firma]['user'] = default_user
                    except NameError:
                        print("Error! User")
                        self.__spo_machine_conf[firma]['sync'] = False
                        self.__spo_machine_conf[firma][
                            'message'] = (
                                f'Fehler! Für die Maschinenanhangs' +
                                f'synchronisation der Firma ' +
                                f'{self.__companies[firma]["name"]} ' +
                                f'Konne kein Benutzer ermittelt werden!')
                if 'pass' in spo_fa_conf:
                    self.__spo_machine_conf[firma]['pass'] = spo_fa_conf['pass']
                else:
                    try:
                        self.__spo_machine_conf[firma]['pass'] = default_pass
                    except NameError:
                        print("Error! Pass")
                        self.__spo_machine_conf[firma]['sync'] = False
                        self.__spo_machine_conf[firma][
                            'message'] = (
                                f'Fehler! Für die Maschinenanhangs' +
                                f'synchronisation der Firma ' +
                                f'{self.__companies[firma]["name"]} ' +
                                f'konne kein Password ermittelt werden!')
                if 'machine_website' in spo_fa_conf:
                    self.__spo_machine_conf[firma]['website'] = (
                        spo_default_conf['machine_website'])
                else:
                    try:
                        self.__spo_machine_conf[firma]['website'] = (
                            default_website)
                    except NameError:
                        print("Error! Website")
                        self.__spo_machine_conf[firma]['sync'] = False
                        self.__spo_machine_conf[firma][
                            'message'] = (
                                f'Fehler! Für die Maschinenanhangs' +
                                f'synchronisation der Firma ' +
                                f'{self.__companies[firma]["name"]} ' +
                                f'Konne keine Website ermittelt werden!')
                if 'machine_library' in spo_default_conf:
                    self.__spo_machine_conf[firma]['library'] = (
                        spo_default_conf['machine_library'])
                else:
                    try:
                        self.__spo_machine_conf[firma]['library'] = (
                            default_library)
                    except NameError:
                        self.__spo_machine_conf[firma]['sync'] = False
                        self.__spo_machine_conf[firma][
                            'message'] = (
                                f'Fehler! Für die Maschinenanhangs' +
                                f'synchronisation der Firma ' +
                                f'{self.__companies[firma]["name"]} ' +
                                f'Konne keine Bibliothek ermittelt werden!')
                # Wenn keine Angaben zu Docksets gemacht werden, werden bei der 
                # Syncronisation, nur Ordner ohne Methadaten erstellt    
                # Kennzeichnet, ob nur Ordner benutzt werden sollen
                self.__spo_machine_conf[firma]['onlyfolder'] = False
                if 'machine_docset' in spo_default_conf:
                    self.__spo_machine_conf[firma]['docset'] = (
                        spo_default_conf['machine_docset'])
                else:
                    try:
                        self.__spo_machine_conf[firma]['docset'] = (
                            default_docset)
                    except NameError:
                        # Nur Ordner nutzen
                        self.__spo_machine_conf[firma]['onlyfolder'] = True
                if 'machine_docset_id' in spo_default_conf:
                    self.__spo_machine_conf[firma]['docset_id'] = (
                        spo_default_conf['machine_docset_id'])
                else:
                    try:
                        self.__spo_machine_conf[firma]['docset_id'] = (
                            default_docset_id)
                    except NameError:
                        # Nur Ordner nutzen
                        self.__spo_machine_conf[firma]['onlyfolder'] = True
            else:
                # Wenn keine gesonderte Conf für Firma in Configdatei
                # Defaultwerte übernehmen
                # Feststellen ob per Default Maschinen synchronisiert 
                # werden sollen
                if default_masch_sync == True:
                    self.__spo_machine_conf[firma]['sync'] = True
                    self.__spo_machine_conf[firma]['message'] = (
                            f'Keine gesonderte Konfiguration für Firma ' +
                            f'{self.__companies[firma]["name"]} ermittelt.\nFür ' +
                            f'die Syncronisation werden Defaultwerte übernommen.')
                    # Feststellen, ob Defaultwerte vollständig
                    if conf_error == False:
                        self.__spo_machine_conf[firma]['user'] = default_user
                        self.__spo_machine_conf[firma]['pass'] = default_pass
                        self.__spo_machine_conf[firma][
                            'website'] = default_website
                        self.__spo_machine_conf[firma][
                            'library'] = default_library
                        # Docset oder Folder
                        if default_onlyfolder == False:
                            self.__spo_machine_conf[firma]['onlyfolder'] = False
                            self.__spo_machine_conf[firma][
                                'docset'] = default_docset
                            self.__spo_machine_conf[firma][
                                'docset_id'] = default_docset_id
                        else:
                            self.__spo_machine_conf[firma]['onlyfolder'] = True
                    else:
                        self.__spo_machine_conf[firma]['sync'] = False
                        self.__spo_machine_conf[firma]['message'] = (
                                f'Keine gesonderte Konfiguration für Firma ' +
                                f'{self.__companies[firma]["name"]} ermittelt.' +
                                f'\Die Defaultwerte sind unvolstandig!' +
                                f'\n{conf_error_msg}')
                else:
                    self.__spo_machine_conf[firma]['sync'] = False
                    self.__spo_machine_conf[firma]['message'] = (
                            f'Keine gesonderte Konfiguration für Firma ' +
                            f'{self.__companies[firma]["name"]} ermittelt.\Die ' +
                            f'Maschinensynchronisation ist per Defaul abgestellt!')

    def get_spo_machine_config(self, firma_id):
        """
        Gibt die Configurationsdaten für die Bibliothek
        der Maschinenanhänge für eine Firma zurück.
        
        Rückgabe Maschinenconfig für FA
        """

        return self.__spo_machine_conf[firma_id]

    def __init_spo_machine_connection(self):
        """
        Initialisiert die Zugriffsobjekte für die
        Bibliothek/en für die Maschinenanhänge
        
        Es wird ein Dictionary mit folgenden Aufbau erstellt:
        
        self.__spo_machine_conn[<ID>]['status']      True/False giebt an, ob ein
                                                    Verbindungs-Objekt zur  
                                                    erfolgreich eingerichtet 
                                                    wurde.
        self.__spo_machine_conn[<ID>]['message']    Text. Giebt die 
                                                    Fehler-/OK-Meldung für 
                                                    das DB-Verbindungs-Objekt
        self.__spo_machine_conn[<ID>]['connection'] Das eigentliche 
                                                    
                                                    DB-Verbindungs-Objekt
        """

        # Config für Maschinen-Bibliotheken ermitteln
        self.__read_spo_machine_config()
        print("---2---")
        # Konfigurationen der einzelnen Firmen abarbeiten 
        # und Connection-Objekte erstellen
        self.__spo_machine_conn = {}
        for conf in self.__spo_machine_conf:
            self.__spo_machine_conn[conf] = {}
            # Connection-Object nur erstellen wenn benötigt!
            print("--> ", self.__spo_machine_conf[conf])
            if self.__spo_machine_conf[conf]['sync'] == True:
                self.__spo_machine_conn[conf]['conn'] = spo_connect.SPOLibrary(
                    self.__spo_machine_conf[conf]['website'], self.__spo_machine_conf[conf]['library'], self.__spo_machine_conf[conf]['user'],
                    self.__spo_machine_conf[conf]['pass'])
                # Connection testen
                result = self.__spo_machine_conn[conf]['conn'].test_connection()
                if result[0] == True:
                    self.__spo_machine_conn[conf]['status'] = True
                    self.__spo_machine_conn[conf]['message'] = 'Verbindung kann aufgebaut werden!'
                else:
                    self.__spo_machine_conn[conf]['status'] = False
                    self.__spo_machine_conn[conf]['message'] = 'Fehler! Verbindung kann mit der Konfiguration nicht aufgebaut werden!'
            else:
                self.__spo_machine_conn[conf]['status'] = False
                self.__spo_machine_conn[conf]['message'] = conf['message']

    def get_spo_machine_connection(self, firmen_id):
        """
        Giebt das SPO-Maschinen-Connection-Objekt für eine Firma zurück.
        
        Rückgabe: Liste erster Wert True/False, zweiter Wert Meldung, dritter Wert das Connection-Object
        """
        if self.__spo_create_mach_conn == True:

            print(len(self.__spo_machine_conn))
            if firmen_id in self.__spo_machine_conn:
                return self.__spo_machine_conn[firmen_id]['status'], self.__spo_machine_conn[firmen_id]['message'], self.__spo_machine_conn[firmen_id]['conn']
            else:
                return False, f'Fehler! Für die Firma mit ID {firmen_id} giebt es keine SPO-Maschinen-verbindung!', None
        else:
            return False, 'Bei der Modulinitialisierung wurde angegeben, dass keine SPO-Connection-Objekte für die Maschinen erstellt werden sollen!', None

    def __read_spo_crm_config(self):
        """
        Liest die Konfigurationen für die SPO-Verbindungen für die Bibliotheken der Verkaufsprojekte (CRM).
        
        Die SPO-Konfigurationen werden als self.__spo_crm_conf[<ID>] ausgegeben
        Dabei entspricht die ID der Firmenid im Oddo. Die ID=0 giebt die Defaultwerte an
        """
        # Defaul-Werte für SPO-Konfugurationen ermitteln
        conf_error = False
        if 'SPO' in self.__config:
            spo_default_conf = self.__config['SPO']
            # Feststellen, ob CRM/Sales-Sync default abgestellt ist
            default_crm_sync = True
            if 'sync_crm' in spo_default_conf:
                if spo_default_conf['sync_crm'] == False:
                    default_crm_sync = False

            # Weiter Default-Werte nur ermitteln, wenn Sync=True
            if default_crm_sync == True:
                if 'user' in spo_default_conf:
                    default_user = spo_default_conf['user']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe eines Benutzers im Abschnit [SPO]!"
                    )
                if 'pass' in spo_default_conf:
                    default_pass = spo_default_conf['pass']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe eines Passwortes im Abschnit [SPO]!"
                    )
                if 'sales_crm_website' in spo_default_conf:
                    default_website = spo_default_conf['sales_crm_website']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe einer CRM-Website im Abschnit [SPO]!"
                    )
                if 'sales_crm_library' in spo_default_conf:
                    default_library = spo_default_conf['sales_crm_library']
                else:
                    conf_error = True
                    conf_error_msg = (
                        "Keine Angabe einer CRM-Bibliothek im Abschnit [SPO]!"
                    )
                default_crm_onlyfolder = False
                if 'crm_docset' in spo_default_conf:
                    default_crm_docset = spo_default_conf['crm_docset']
                else:
                    default_crm_onlyfolder = True
                if 'crm_docset_id' in spo_default_conf:
                    default_crm_docset_id = spo_default_conf['crm_docset_id']
                else:
                    default_crm_onlyfolder = True
                default_sales_onlyfolder = False
                if 'sales_docset' in spo_default_conf:
                    default_sales_docset = spo_default_conf['sales_docset']
                else:
                    default_sales_onlyfolder = True
                if 'sales_docset_id' in spo_default_conf:
                    default_sales_docset_id = spo_default_conf['sales_docset_id']
                else:
                    default_sales_onlyfolder = True

        # Konfigurationen für die einzelnen Firmen ermitteln
        self.__spo_crm_conf = {}
        for firma in self.__companies:
            self.__spo_crm_conf[firma] = {}

            # Feststellen, ob gesonderte Konfiguration
            # für die entsprechende FA vorhanden
            co_sec = f'SPO_{firma}'
            if co_sec in self.__config:
                # Wenn gesonderte Konfiguration vorhanden
                spo_fa_conf = self.__config[co_sec]
                self.__spo_crm_conf[firma]['sync'] = True
                self.__spo_crm_conf[firma]['message'] = f'Konfiguration für Firma {self.__companies[firma]["name"]} ermittelt'
                # Feststellen, ob CRManhänge nict syncronisiert werden sollen
                if 'sync_crm' in spo_fa_conf:
                    if spo_fa_conf['sync_crm'] == False:
                        self.__spo_crm_conf[firma]['sync'] = False
                        self.__spo_crm_conf[firma][
                            'message'] = f'Laut Konfiguration soll für Firma {self.__companies[firma]["name"]} keine CRM-Anhangssynchronisation durchgeführt werden.'
                if 'user' in spo_fa_conf:
                    self.__spo_crm_conf[firma]['user'] = spo_fa_conf['user']
                else:
                    try:
                        self.__spo_crm_conf[firma]['user'] = default_user
                    except NameError:
                        self.__spo_crm_conf[firma]['sync'] = False
                        self.__spo_crm_conf[firma][
                            'message'] = f'Fehler! Für die CRM-Anhangssynchronisation der Firma {self.__companies[firma]["name"]} Konne kein Benutzer ermittelt werden!'
                if 'pass' in spo_fa_conf:
                    self.__spo_crm_conf[firma]['pass'] = spo_fa_conf['pass']
                else:
                    try:
                        self.__spo_crm_conf[firma]['pass'] = default_pass
                    except NameError:
                        self.__spo_crm_conf[firma]['sync'] = False
                        self.__spo_crm_conf[firma][
                            'message'] = f'Fehler! Für die CRM-Anhangssynchronisation der Firma {self.__companies[firma]["name"]} Konne kein Password ermittelt werden!'
                if 'sales_crm_website' in spo_fa_conf:
                    self.__spo_crm_conf[firma]['website'] = spo_default_conf['sales_crm_website']
                else:
                    try:
                        self.__spo_crm_conf[firma]['website'] = default_website
                    except NameError:
                        self.__spo_crm_conf[firma]['sync'] = False
                        self.__spo_crm_conf[firma][
                            'message'] = f'Fehler! Für die CRM-Anhangssynchronisation der Firma {self.__companies[firma]["name"]} Konne keine Website ermittelt werden!'
                if 'crm_library' in spo_default_conf:
                    self.__spo_crm_conf[firma]['library'] = spo_default_conf['crm_library']
                else:
                    try:
                        self.__spo_crm_conf[firma]['library'] = default_library
                    except NameError:
                        self.__spo_crm_conf[firma]['sync'] = False
                        self.__spo_crm_conf[firma][
                            'message'] = f'Fehler! Für die CRM-Anhangssynchronisation der Firma {self.__companies[firma]["name"]} Konne keine Bibliothek ermittelt werden!'
                # Wenn keine Angaben zu CRM-Docksets gemacht werden, werden bei der Syncronisation, nur Ordner ohne Methadaten erstellt
                # Kennzeichnet, ob nur Ordner für CRM benutzt werden sollen
                self.__spo_crm_conf[firma]['crm_onlyfolder'] = False
                if 'crm_docset' in spo_default_conf:
                    self.__spo_crm_conf[firma]['crm_docset'] = spo_default_conf['crm_docset']
                else:
                    try:
                        self.__spo_crm_conf[firma]['crm_docset'] = default_crm_docset
                    except NameError:
                        # Nur Ordner nutzen
                        self.__spo_crm_conf[firma]['crm_onlyfolder'] = True
                if 'crm_docset_id' in spo_default_conf:
                    self.__spo_crm_conf[firma]['crm_docset_id'] = spo_default_conf['crm_docset_id']
                else:
                    try:
                        self.__spo_crm_conf[firma]['crm_docset_id'] = default_crm_docset_id
                    except NameError:
                        # Nur Ordner nutzen
                        self.__spo_crm_conf[firma]['crm_onlyfolder'] = True
                # Wenn keine Angaben zu Sales-Docksets gemacht werden, werden bei der Syncronisation, nur Ordner ohne Methadaten erstellt
                # Kennzeichnet, ob nur Ordner für Sales benutzt werden sollen
                self.__spo_crm_conf[firma]['sales_onlyfolder'] = False
                if 'sales_docset' in spo_default_conf:
                    self.__spo_crm_conf[firma]['sales_docset'] = spo_default_conf['sales_docset']
                else:
                    try:
                        self.__spo_crm_conf[firma]['sales_docset'] = default_sales_docset
                    except NameError:
                        # Nur Ordner nutzen
                        self.__spo_crm_conf[firma]['sales_onlyfolder'] = True
                if 'crm_docset_id' in spo_default_conf:
                    self.__spo_crm_conf[firma]['sales_docset_id'] = spo_default_conf['sales_docset_id']
                else:
                    try:
                        self.__spo_crm_conf[firma]['sales_docset_id'] = default_sales_docset_id
                    except NameError:
                        # Nur Ordner nutzen
                        self.__spo_crm_conf[firma]['sales_onlyfolder'] = True
            else:
                # Wenn keine gesonderte Conf für die FA in Confdatei,
                # Defaultwerte übernehmen
                # Feststellen, ob per Default CRM und Sales 
                # syncronisiert werden sollen
                if default_crm_sync == True:
                    self.__spo_crm_conf[firma]['sync'] = True
                    self.__spo_crm_conf[firma]['message'] = (
                            f'Keine gesonderte Konfiguration für Firma ' +
                            f'{self.__companies[firma]["name"]} ermittel.\n Für ' +
                            f'die Syncronisation werden Defaultwerte übernommen.')
                    # Feststellen, ob Default-Werte volständig
                    if conf_error == False:
                        self.__spo_crm_conf[firma]['user'] = default_user
                        self.__spo_crm_conf[firma]['pass'] = default_pass
                        self.__spo_crm_conf[firma]['website'] = default_website
                        self.__spo_crm_conf[firma]['library'] = default_library
                        # Docset oder Folder für CRM
                        if default_crm_onlyfolder == False:
                            self.__spo_crm_conf[firma]['crm_onlyfolder'] = False
                            self.__spo_crm_conf[firma][
                                'crm_docset'] = default_crm_docset
                            self.__spo_crm_conf[firma][
                                'crm_docset_id'] = default_crm_docset_id
                        else:
                            self.__spo_crm_conf[firma]['crm_onlyfolder'] = True
                        # Docset oder Folder für Sales
                        if default_sales_onlyfolder == False:
                            self.__spo_crm_conf[firma]['sales_onlyfolder'] = False
                            self.__spo_crm_conf[firma][
                                'sales_docset'] = default_sales_docset
                            self.__spo_crm_conf[firma][
                                'sales_docset_id'] = default_sales_docset_id
                        else:
                            self.__spo_crm_conf[firma]['sales_onlyfolder'] = True
                else:
                    self.__spo_crm_conf[firma]['sync'] = False
                    self.__spo_crm_conf[firma]['message'] = (
                            f'Keine gesonderte Konfiguration für Firma ' +
                            f'{self.__companies[firma]["name"]} ermittel.\n Die ' +
                            f'CRM-Syncronisation ist per Default abgestellt!.')

    def get_spo_crm_config(self, firmen_id):
        """
        Gibt die Configurationsdaten für die Bibliothek
        der Maschinenanhänge für eine Firma zurück.
        
        Rückgabe CRM/Sales-Conf für FA
        """

        return self.__spo_crm_conf[firmen_id]

    def __init_spo_crm_connection(self):
        """
        Initialisiert die Zugriffsobjekte für die
        Bibliothek/en für die Verkaufsprojekte
        
        Es wird ein Dictionary mit folgenden Aufbau erstellt:
        
        self.__spo_crm_conn[<ID>]['status']      True/False giebt an, ob ein
                                                Verbindungs-Objekt zur  erfolgreich 
                                                eingerichtet wurde.
        self.__spo_crm_conn[<ID>]['message']    Text. Giebt die Fehler-/OK-Meldung für 
                                                das DB-Verbindungs-Objekt
        self.__spo_crm_conn[<ID>]['connection'] Das eigentliche DB-Verbindungs-Objekt
        """

        # Config für CRM-Bibliotheken ermitteln
        self.__read_spo_crm_config()
        # Konfigurationen der einzelnen Firmen abarbeiten
        # und Connection-Objekte erstellen
        self.__spo_crm_conn = {}
        for conf in self.__spo_crm_conf:
            self.__spo_crm_conn[conf] = {}
            # Connection-Object nur erstellen wenn benötigt!
            if self.__spo_crm_conf[conf]['sync'] == True:
                self.__spo_crm_conn[conf]['conn'] = spo_connect.SPOLibrary(
                    self.__spo_crm_conf[conf]['website'], self.__spo_crm_conf[conf]['library'], self.__spo_crm_conf[conf]['user'],
                    self.__spo_crm_conf[conf]['pass'])
                # Connection testen
                result = self.__spo_crm_conn[conf]['conn'].test_connection()
                if result[0] == True:
                    self.__spo_crm_conn[conf]['status'] = True
                    self.__spo_crm_conn[conf]['message'] = 'Verbindung kann aufgebaut werden!'
                else:
                    self.__spo_crm_conn[conf]['status'] = False
                    self.__spo_crm_conn[conf]['message'] = 'Fehler! Verbindung kann mit der Konfiguration nicht aufgebaut werden!'
            else:
                self.__spo_crm_conn[conf]['status'] = False
                self.__spo_crm_conn[conf]['message'] = conf['message']

    def get_spo_crm_connection(self, firmen_id):
        """
        Giebt das SPO-Verkaufsprojekte-Connection-Objekt für eine Firma zurück.
        
        Rückgabe: Liste erster Wert True/False, zweiter Wert Meldung, dritter Wert das Connection-Object
        """
        if self.__spo_create_crm_conn == True:
            if firmen_id in self.__spo_crm_conn:
                print(self.__spo_crm_conn[firmen_id])
                return self.__spo_crm_conn[firmen_id]['status'], self.__spo_crm_conn[firmen_id]['message'], self.__spo_crm_conn[firmen_id]['conn']
            else:
                return False, f'Fehler! Für die Firma mit ID {firmen_id} giebt es keine SPO-Verkaufsprojekte-verbindung!'
        else:
            return False, 'Bei der Modulinitialisierung wurde angegeben, dass keine SPO-Connection-Objekte für die Verkaufsprojekte erstellt werden sollen!'
