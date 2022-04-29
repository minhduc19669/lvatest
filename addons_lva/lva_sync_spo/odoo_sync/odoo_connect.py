""" 
Diese Datei ist Bestandteil der Dateisyncronisation
zwischen odoo und SPO. Dabei geht es darum, zu einem
entsprechenden Objekt im Odoo ein entsprechendes Docsets
anzulegen und vorhandene dateien dort abzulegen.
Sie stellt die Objekte und Funktionen bereit, 
die für die Verbindung mit der odoo-db benötigt werden.
"""

import psycopg2
import psycopg2.extras
from psycopg2 import Error
from sshtunnel import SSHTunnelForwarder


class OdooDB:
    """Diese Klasse dient dem Zugriff auf eine odoo-datenbank"""

    def __init__(self, db_host, db_user, db_pass, db_db, db_port=5432):
        """ 
        Hinterlegt die Grunddaten für den Zugriff auf eine Odoo-DB

            db_host  - Datenbankserver
            db_user  - Datenbankbenutzer
            db_pass  - Password
            db_db    - Name der Datenbank
            db_port  - Port des DB-Srvers (optional)
        """
        # Zugangsdaten für den DB-Server
        self._host = db_host
        self._user = db_user
        self.__pass = db_pass
        self._db = db_db
        self._port = db_port

        # Zugangsdaten für den ssh-Tunnel (nur bei Bedarf)
        self.__use_ssh_tunnel = False
        # self._ssh_host=''
        # self._ssh_port=0
        # self._ssh_user=''
        # self.__ssh_pass=''
        # self._ssh_remote_host=''
        # self._ssh_remote_port=0
        # self._ssh_local_host=''
        # self._ssh_local_port=0

    def ssh_tunnel_init(self, ssh_host, ssh_user, ssh_pass, ssh_port=22):
        """
        Initialisiert einen SSH-Tunnel. Dieser wird gegebenenfalls benötigt, 
        wenn der Datenbankserver nicht direkt erreichbar ist.
            ssh_host            = Hoatadresse des SSH-Servers    
            ssh_port (opt.)     = SSH-Portnummer auf SSH-Server
            ssh_user            = Benutzer auf SSH-Server
            ssh_pass            = Passwort auf SSH-Server
            remote_host         = Hostadresse des zu tunnelden Dienstes
            remote_port (opt.)  = Portnummer des zu tunnelden Dienstes
            local_host (opt.)   = Locale Adresse an die der Tunnel gebunden 
                                  werden soll
            local_port (opt.)   = Portnummer an der der Tunnel gebunden 
                                  werden soll

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung
        """
        # Den tunnel-remote-host auf Datenban-host setzen
        remote_host = self._host
        # Datenbankhost auf für Tunnelung auf localhost setzen
        self._host = '127.0.0.1'
        print(ssh_host, remote_host, self._host)

        self.__ssh_tunnel = SSHTunnelForwarder(
            ssh_host, ssh_port=ssh_port,
            ssh_username=ssh_user,
            ssh_password=ssh_pass,
            remote_bind_address=(remote_host, self._port),
            local_bind_address=(self._host, self._port)
        )

        print(1)
        try:
            self.__ssh_tunnel.start()
            self.__ssh_tunnel.stop()
            self.__use_ssh_tunnel = True
            return True, (
                f'Kann SSH-Tunnel zum Server {ssh_host}:{ssh_port} aufbauen.')
        except:
            return False, (
                    f'Kann keinen SSH-Tunnel zum Server ' +
                    f'{ssh_host}:{ssh_port} aufbauen!')

    def test_connect(self):
        """ testet die Verbindung zur Datenbank """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, (
                            'Kann den SSH-Tunnel für die ' +
                            'Datenbanverbindung nicht aufbauen!')
            # Datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, port=self._port, dbname=self._db, user=self._user,
                password=self.__pass)
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()

            return True, (
                    f'Eine Verbindung zur Datenbank {self._db} auf dem Server ' +
                    f'{self._host} konte aufgebaut werden.')
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db} ' +
                    f'auf {self._host} aufgebaut werden!')

    def get_company_list(self):
        """
        Liest die Firmen aus der Odoo-Datenbank aus und giebt dies zurück.

            Rückgabe: Liste, erter Wert True/False, zweiter Wert das Diktionary 
                      mit den Zeilen oder eine Fehlermeldung
        """
        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(host=self._host, dbname=self._db,
                                    user=self._user, password=self.__pass,
                                    port=self._port,
                                    cursor_factory=
                                    psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db}' +
                    f'auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = "select id, name from res_company;"
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            return True, rows
        except:
            return False, (
                'Die Firmenliste konnte nicht erfogreich abgefragt werden!')

    def get_machine_list(self, company_id=None, rowlimit=0, offset=0):
        """
        Liest die Maschinenliste mit einer Select Anweisung aus und
        giebt diese als Dictionary zurück. Die beötigten, verküpften Daten
        werden mittels Left Join mit abgefragt.
            company_id (optional)  = giebt optional die Firmenid zur 
                                     Filterung an.
            rowlimmit (optional)   = giebt an, wie viele Zeilen abgefragt werden
                                     sollen. Bei 0 oder Nichtangabe werden alle 
                                     Zeilen ausgelesen.
            offset (optional)      = Giebt an, ab welcher Zeile ausgegeben 
                                     werden soll. Bei 0 oder Nichtangabe wird 
                                     ab der ersten Zeile ausgelesen.
            
            Rückgabe: Liste, erter Wert True/False, zweiter Wert Meldung, 
                      dritter Wert das Diktionary mit den Zeilen oder eine 
                      Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db} auf' +
                    f'{self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        # Filter nach Firmenid
        where_company = ""
        if company_id != None:
            where_company = f' where firma.id = {company_id} '
        request = f"""
                select
                    ms.id,
                    ms.mx_number as mxnr,
                    ms.categ_id as msartid,
                    cat.name as msart,
                    ms.sub_categ_id as mssubartid,
                    subcat.name as mssubart,
                    ms.manufacturer as hstid,
                    hst.name as hst,
                    ms.model as typeid,
                    md.name as type, 
                    ms.chassis_number as fgnr,
                    ms.license_plate as kenz,
                    ms.new_old_short as gebr,
                    ms.year_of_constr as bj,
                    ms.date_of_first_reg as ezl,
                    ms.client as firmaid,
                    firma.name as firma
                from  machine_list_machine as ms 
                    left join machine_list_category as cat on ms.categ_id=cat.id 
                    left join machine_list_sub_category as subcat on 
                        ms.sub_categ_id=subcat.id
                    left join machine_list_manufacturer as hst on 
                        ms.manufacturer=hst.id
                    left join machine_list_model as md on ms.model=md.id
                    left join res_company as firma on 
                        ms.client = firma.id{where_company}
                order by ms.id
                """
        # Limit und Offset    
        if rowlimit != 0:
            request = request + f'\nlimit {rowlimit} offset {offset}'
        request = request + ";"
        # print(request)
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            return True, 'Maschinen ausgelesen', rows
        except:
            return False, (
                'Die Maschinenliste konnte nicht erfogreich abgefragt werden!',
                None)

    def get_machine_attached_folder_link(self, machine_id):
        """
        Ermittelt, ob ein Eintrag '__(SPO)_Ordner' für die Maschine
        mit der id machine_id giebt.
            machine_id  = Odoo_Id der Maschine

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Liste mit den 
                      Eintragsdaten oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db}' +
                    f'auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                select * from ir_attachment 
                    where 
                        name = '__(SPO)_Ordner' and
                        res_model = 'machine.list.machine' and
                        res_id = {machine_id};
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows
            else:
                return False, (
                        f'Keinen Ordner-Attachment-Eintrag für die Maschine' +
                        f' {machine_id} gefunden!')
        except:
            return False, (
                'Die Attachmentliste konnte nicht erfogreich abgefragt werden!')

    def get_machine_attachment_file_links(
            self, machine_id, attachment_type=None):
        """
        Ermittelt die Dateianhänge für eine Maschine.
        (Die Eintrage für die Ordnerlinks werden ausgelassen!)
            machine_id      = Odoo_Id der Maschine
            attachment_type = Type des Anhangs
                                binary: local auf dem Odoo-Server gespeichert
                                url:    als Link eingetragen
                                keine Angabe: alle Einträge

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Liste der 
                      Einträge oder Fehlermeldung.
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                select * from ir_attachment 
                    where 
                        name != '__(SPO)_Ordner' and
                        res_model = 'machine.list.machine' and
                        res_id = {machine_id}
                """
        if attachment_type != None:
            request = request + f" and type = '{attachment_type}'"

        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows
            else:
                return False, (f'Keinen Attachment-Einträge für die Maschine ' +
                               f'{machine_id} gefunden!')
        except:
            return False, (
                'Die Maschinenliste konnte nicht erfogreich abgefragt werden!')

    def create_machine_attachment_folder_link(
            self, machine_id, machine_name, link_url, firma_id, user_id=1):
        """
        Erstellt einen Ordner-Attachment-Link für eine Maschine.
            machine_id      = Odoo_Id der Maschine
            machine_name    = Name der Maschine in Form 
                              '[<mxnr>] <hersteller> <type>'
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                insert into ir_attachment (
                    name, 
                    datas_fname,
                    res_name,
                    res_model,
                    res_model_name,
                    res_id,
                    company_id,
                    type,
                    url,
                    mimetype,
                    active,
                    create_uid,
                    create_date,
                    write_uid,
                    write_date)
                values (
                    '__(SPO)_Ordner',
                    '__(SPO)_Ordner',
                    '{machine_name}',
                    'machine.list.machine',
                    'Machine list machine',
                    {machine_id},
                    {firma_id},
                    'url',
                    '{link_url}',
                    'application/spo.folder',
                    true,
                    {user_id},
                    now(),
                    {user_id},
                    now())
                returning *;
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                print(rows[0])
                return True, rows[0]['id']
            else:
                return False, (f'Konte keinen Folder-Attachment-Eintrag für ' +
                               f'die Maschine {machine_id} erstellen!')
        except:
            return False, (f'Konte keinen Folder-Attachment-Eintrag für die ' +
                           f'Maschine {machine_id} erstellen!')

    def create_machine_attachment_file_link(
            self, machine_id, machine_name, file_name, link_url, firma_id, mimetype, user_id=1):
        """
        Erstellt einen Ordner-Attachment-Link für eine Maschine.
            machine_id      = Odoo_Id der Maschine
            machine_name    = Name der Maschine im Odoo 
            file_name       = Name der Datei im SPO.
                              Im Link wird dem Namen noch "_(SPO)_" 
                              vorangestellt. Dies kennzeichnet,
                              dass sich der Anhang im Sharepoint befindet.
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            mimetype        = Mimetype für den Anhang z.B. "application/pdf" 
                              o. "image/png"
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass,
                port=self._port, cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                insert into ir_attachment (
                    name, 
                    datas_fname,
                    res_name,
                    res_model,
                    res_model_name,
                    res_id,
                    company_id,
                    type,
                    url,
                    mimetype,
                    active,
                    create_uid,
                    create_date,
                    write_uid,
                    write_date)
                values (
                    '_(SPO)_{file_name}',
                    '{file_name}',
                    '{machine_name}',
                    'machine.list.machine',
                    'Machine list machine',
                    {machine_id},
                    {firma_id},
                    'url',
                    '{link_url}',
                    '{mimetype}',
                    true,
                    {user_id},
                    now(),
                    {user_id},
                    now())
                returning id;
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte keinen File-Attachment-Eintrag für ' +
                               f'die Maschine {machine_id} erstellen!')
        except:
            return False, (f'Konte keinen File-Attachment-Eintrag für ' +
                           f'die Maschine {machine_id} erstellen!')

    def update_machine_attachment_folder_link(
            self, machine_id, machine_name, link_url,
            firma_id, user_id=1, attachment_id=None):
        """
        Ändert einen Ordner-Attachment-Link für eine Maschine.
        Die Filterung erfolgt entweder über die res_id="machine_id" und 
        res_model oder über die id="attachment_id".
        !! Achtung !! ist die "attachment_id" angegeben wird nur 
        über sie gefiltert !!
        Es werden die Felder res_name, company_id und url gesetzt. 
            machine_id      = Odoo_Id der Maschine
            machine_name    = Name der Maschine in Form 
                              '[<mxnr>] <hersteller> <type>'
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!
            attachment_id (opt.)    = Odoo-ID des Anhangeintrags

            Rückgabe: Liste, erster Wert True/False, Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                update ir_attachment 
                set
                    res_name = '{machine_name}',
                    company_id  = '{firma_id}',
                    url = '{link_url}',
                    active = true,
                    write_uid = '{user_id}',
                    write_date = now()
                where
                    res_model = 'machine.list.machine' and
                    res_id = {machine_id} and
                    name = '__(SPO)_Ordner'
                returning id;
                """
        if attachment_id != None:
            request = f"""
                    update ir_attachment 
                    set
                        res_name = '{machine_name}',
                        company_id  = '{firma_id}',
                        res_id = {machine_id},
                        url = '{link_url}',
                        active = true,
                        write_uid = '{user_id}',
                        write_date = now()
                    where
                        id = {attachment_id}
                    returning *;
                    """

            # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte den Folder-Attachment-Eintrag für ' +
                               f'die Maschine {machine_id} nicht ändern!')
        except:
            return False, (f'Konte den Folder-Attachment-Eintrag für die ' +
                           f'Maschine {machine_id} nicht ändern!')

    def update_machine_attachment_file_link(
            self, attachment_id, machine_id, file_name, link_url,
            firma_id, mimetype, user_id=1):
        """
        Aktualisiert einen Ordner-Attachment-Link für eine Maschine.
        !! Achtung auch Anhangeinträge vom Type "binary" werden in einen 
        Link umgewandelt !!
            attachment_id   = Odoo-Id des Anhangeintrags
                              Die Filterung erfolgt nur über die ID!! 
            machine_id      = Odoo_Id der Maschine
            file_name       = Name der Datei im SPO. Im Link wird dem Namen 
                              nach "_(SPO)_" vorangestellt. Dies kennzeichnet,
                              dass sich der Anhang im Sharepoint befindet.
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            mimetype        = Mimetype für den Anhang z.B. "application/pdf" 
                              o. "image/png"
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                    update ir_attachment 
                    set
                        name = '_(SPO)_{file_name}',
                        datas_fname = '{file_name}',
                        company_id  = '{firma_id}',
                        res_id = {machine_id},
                        url = '{link_url}',
                        type = 'url',
                        active = true,
                        write_uid = '{user_id}',
                        write_date = now(),
                        mimetype='{mimetype}'
                    where
                        id = {attachment_id}
                    returning *;
                    """
        # print(request)
        # Abfrage ausführen
        cursor = conn.cursor()
        # try:
        cursor.execute(request)
        conn.commit()
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        # Bei Bedarf, SSH-Tunnel schließen
        if self.__use_ssh_tunnel == True:
            self.__ssh_tunnel.stop()
        if len(rows) > 0:
            # print(rows[0])
            return True, rows[0]['id']
        else:
            return False, (f'Konte den File-Attachment-Eintrag für die ' +
                           f'Maschine {machine_id} nicht ändern!')

    def get_crm_list(self, company_id=None, rowlimit=0, offset=0):
        """
        Liest die CRM-Liste mit einer Select Anweisung aus und
        giebt diese als Dictionary zurück. Die beötigten, verküpften Daten
        werden mittels Left Join mit abgefragt.
            company_id (optional)  = giebt optional die Firmenid zur 
                                     Filterung an.
            rowlimmit (optional)   = giebt an, wie viele Zeilen abgefragt werden
                                     sollen. Bei 0 oder Nichtangabe werden alle 
                                     Zeilen ausgelesen.
            offset (optional)      = Giebt an, ab welcher Zeile ausgegeben 
                                     werden soll. Bei 0 oder Nichtangabe wird 
                                     ab der ersten Zeile ausgelesen.
            
            Rückgabe: Liste, erter Wert True/False, zweiter Wert Meldung, 
                      dritter Wert das Diktionary mit den Zeilen oder eine 
                      Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db} auf' +
                    f'{self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        # Filter nach Firmenid
        where_company = ""
        if company_id != None:
            where_company = f' where firma.id = {company_id} '
        request = f"""
                select 
                    crm.id,
                    crm.name,
                    crm.partner_id,
                    (case 
                        when crm.partner_name > '' and crm.contact_name > '' then concat(crm.contact_name ,', ',crm.partner_name)
                        when crm.partner_name > '' then crm.partner_name 
                        else crm.contact_name end)	as partner,
                    crm.stage_id ,
                    stage."name" as status,
                    crm.user_id ,
                    users.login,
                    crm.company_id as firmaid,
                    firma.name as firma
                from crm_lead as crm
                    left join crm_stage as stage on crm.stage_id = stage.id 
                    left join res_users as users on crm.user_id =users.id 
                    left join res_company as firma on crm.company_id =firma.id {where_company}
                order by crm.id
                """
        # Limit und Offset
        if rowlimit != 0:
            request = request + f'\nlimit {rowlimit} offset {offset}'
        request = request + ";"
        # print(request)
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            return True, 'CRM-Liste ausgelesen', rows
        except:
            return False, (
                'Die CRM-Liste konnte nicht erfogreich abgefragt werden!',
                None)

    def get_crm_attachment_file_links(
            self, crm_id, attachment_type=None):
        """
        Ermittelt die Dateianhänge für eine Maschine.
        (Die Eintrage für die Ordnerlinks werden ausgelassen!)
            crm_id          = Odoo_Id der Maschine
            attachment_type = Type des Anhangs
                                binary: local auf dem Odoo-Server gespeichert
                                url:    als Link eingetragen
                                keine Angabe: alle Einträge

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Liste der 
                      Einträge oder Fehlermeldung.
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                select * from ir_attachment 
                    where 
                        name != '__(SPO)_Ordner' and
                        res_model = 'crm.lead' and
                        res_id = {crm_id}
                """
        if attachment_type != None:
            request = request + f" and type = '{attachment_type}'"

        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows
            else:
                return False, (f'Keinen Attachment-Einträge für das CRM ' +
                               f'{crm_id} gefunden!')
        except:
            return False, (
                'Die Attachmentliste konnte nicht erfogreich abgefragt werden!')

    def get_crm_attachment_folder_link(self, crm_id):
        """
        Ermittelt, ob ein Eintrag '__(SPO)_Ordner' für das CRM
        mit der id crm_id giebt.
            crm_id  = Odoo_Id des CRM

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Liste mit den 
                      Eintragsdaten oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db}' +
                    f'auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                select * from ir_attachment 
                    where 
                        res_model = 'crm.lead' and
                        res_id = {crm_id};
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows
            else:
                return False, (
                        f'Keinen Ordner-Attachment-Eintrag für das CRM' +
                        f' {crm_id} gefunden!')
        except:
            return False, (
                'Die Attachmentliste konnte nicht erfogreich abgefragt werden!')

    def create_crm_attachment_folder_link(
            self, crm_id, crm_name, link_url, firma_id, user_id=1):
        """
        Erstellt einen Ordner-Attachment-Link für ein CRM.
            crm_id          = Odoo_Id des CRM
            crm_name        = Name des CRM in Form 
                              '[<mxnr>] <hersteller> <type>'
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                insert into ir_attachment (
                    name, 
                    datas_fname,
                    res_name,
                    res_model,
                    res_model_name,
                    res_id,
                    company_id,
                    type,
                    url,
                    mimetype,
                    active,
                    create_uid,
                    create_date,
                    write_uid,
                    write_date)
                values (
                    '__(SPO)_Ordner',
                    '__(SPO)_Ordner',
                    '{crm_name}',
                    'crm.lead',
                    'Lead/Opportunity',
                    {crm_id},
                    {firma_id},
                    'url',
                    '{link_url}',
                    'application/spo.folder',
                    true,
                    {user_id},
                    now(),
                    {user_id},
                    now())
                returning *;
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            print('????  ', len(rows))
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte keinen Folder-Attachment-Eintrag für ' +
                               f'das CRM {crm_id} erstellen!')
        except:
            return False, (f'Konte keinen Folder-Attachment-Eintrag für das ' +
                           f'CRM {crm_id} erstellen!')

    def create_crm_attachment_file_link(
            self, crm_id, crm_name, file_name, link_url, firma_id, mimetype, user_id=1):
        """
        Erstellt einen Ordner-Attachment-Link für eine Maschine.
            crm_id          = Odoo_Id des CRM
            crm_name        = Names des CRM im Odoo
            file_name       = Name der Datei im SPO.
                              Im Link wird dem Namen noch "_(SPO)_" 
                              vorangestellt. Dies kennzeichnet,
                              dass sich der Anhang im Sharepoint befindet.
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            mimetype        = Mimetype für den Anhang z.B. "application/pdf" 
                              o. "image/png"
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass,
                port=self._port, cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                insert into ir_attachment (
                    name, 
                    datas_fname,
                    res_name,
                    res_model,
                    res_model_name,
                    res_id,
                    company_id,
                    type,
                    url,
                    mimetype,
                    active,
                    create_uid,
                    create_date,
                    write_uid,
                    write_date)
                values (
                    '_(SPO)_{file_name}',
                    '{file_name}',
                    '{crm_name}',
                    'crm.lead',
                    'Lead/Opportunity',
                    {crm_id},
                    {firma_id},
                    'url',
                    '{link_url}',
                    '{mimetype}',
                    true,
                    {user_id},
                    now(),
                    {user_id},
                    now())
                returning id;
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte keinen File-Attachment-Eintrag für ' +
                               f'das CRM {crm_id} erstellen!')
        except:
            return False, (f'Konte keinen File-Attachment-Eintrag für ' +
                           f'das CRM {crm_id} erstellen!')

    def update_crm_attachment_folder_link(
            self, crm_id, crm_name, link_url,
            firma_id, user_id=1, attachment_id=None):
        """
        Ändert einen Ordner-Attachment-Link für ein CRM.
        Die Filterung erfolgt entweder über die res_id="crm_id" und 
        res_model oder über die id="attachment_id".
        !! Achtung !! ist die "attachment_id" angegeben wird nur 
        über sie gefiltert !!
        Es werden die Felder res_name, company_id und url gesetzt. 
            crm_id          = Odoo_Id des CRM
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!
            attachment_id (opt.)    = Odoo-ID des Anhangeintrags

            Rückgabe: Liste, erster Wert True/False, Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                update ir_attachment 
                set
                    res_name = '{crm_name}',
                    company_id  = '{firma_id}',
                    url = '{link_url}',
                    active = true,
                    write_uid = '{user_id}',
                    write_date = now()
                where
                    res_model = 'crm.lead' and
                    res_id = {crm_id}
                returning id;
                """
        if attachment_id != None:
            request = f"""
                    update ir_attachment 
                    set
                        company_id  = {firma_id},
                        res_id = {crm_id},
                        url = '{link_url}',
                        write_uid = {user_id},
                        write_date = now()
                    where
                        id = {attachment_id}
                    """

        # Abfrage ausführen
        cursor = conn.cursor()
        import sys
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte den Folder-Attachment-Eintrag für ' +
                               f'das CRM {crm_id} nicht ändern!')
        except:
            return False, (f'Konte den Folder-Attachment-Eintrag für das ' +
                           f'CRM {crm_id} nicht ändern!')

    def update_crm_attachment_file_link(
            self, attachment_id, crm_id, file_name, link_url,
            firma_id, mimetype, user_id=1):
        """
        Aktualisiert einen Ordner-Attachment-Link für eine Maschine.
        !! Achtung auch Anhangeinträge vom Type "binary" werden in einen 
        Link umgewandelt !!
            attachment_id   = Odoo-Id des Anhangeintrags
                              Die Filterung erfolgt nur über die ID!! 
            crm_id          = Odoo_Id des CRM
            file_name       = Name der Datei im SPO. Im Link wird dem Namen 
                              nach "_(SPO)_" vorangestellt. Dies kennzeichnet,
                              dass sich der Anhang im Sharepoint befindet.
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            mimetype        = Mimetype für den Anhang z.B. "application/pdf" 
                              o. "image/png"
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                    update ir_attachment 
                    set
                        name = '_(SPO)_{file_name}',
                        datas_fname = '{file_name}',
                        company_id  = '{firma_id}',
                        res_id = {crm_id},
                        url = '{link_url}',
                        type = 'url',
                        active = true,
                        write_uid = '{user_id}',
                        write_date = now(),
                        mimetype='{mimetype}'
                    where
                        id = {attachment_id}
                    returning *;
                    """
        # print(request)
        # Abfrage ausführen
        cursor = conn.cursor()
        # try:
        cursor.execute(request)
        conn.commit()
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        # Bei Bedarf, SSH-Tunnel schließen
        if self.__use_ssh_tunnel == True:
            self.__ssh_tunnel.stop()
        if len(rows) > 0:
            # print(rows[0])
            return True, rows[0]['id']
        else:
            return False, (f'Konte den File-Attachment-Eintrag für das ' +
                           f'CRM {crm_id} nicht ändern!')

    def get_sales_list(self, company_id=None, crm_id=None, rowlimit=0, offset=0):
        """
        Liest die Sales-Liste mit einer Select Anweisung aus und
        giebt diese als Dictionary zurück. Die beötigten, verküpften Daten
        werden mittels Left Join mit abgefragt.
            company_id (optional)  = giebt optional die Firmenid zur 
                                     Filterung an.
            crm_id (optional)      = Die ID des CRM, dem der Sale zugeordnet ist.
                                     Bei None wird nicht nach crm_id gefiltert.
                                     Bei 'Null' (caseunsensitiv) werden die
                                     Einträge ohne Zuordnung gefiltert.
                                     Bei Angabe einer ID wird nach dieser gefiltert. 
            rowlimmit (optional)   = giebt an, wie viele Zeilen abgefragt werden
                                     sollen. Bei 0 oder Nichtangabe werden alle 
                                     Zeilen ausgelesen.
            offset (optional)      = Giebt an, ab welcher Zeile ausgegeben 
                                     werden soll. Bei 0 oder Nichtangabe wird 
                                     ab der ersten Zeile ausgelesen.
            
            Rückgabe: Liste, erter Wert True/False, zweiter Wert Meldung, 
                      dritter Wert das Diktionary mit den Zeilen oder eine 
                      Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!'), None
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db} auf' +
                    f'{self._host} aufgebaut werden!'), None

        # Select-Abfrage für die Liste
        # Filter nach Firmenid
        where_company = ""
        if company_id != None:
            where_company = f'sale.company_id = {company_id}'
        # Filter nach CRM-ID    
        where_crm = ""
        if crm_id != None:
            if type(crm_id) is str:
                if crm_id.lower() == 'null':
                    where_crm = 'opportunity_id is null'
            if type(crm_id) is int:
                where_crm = f'opportunity_id = {crm_id}'
        # WHERE-Klausel erstellen
        where_clause = ''
        if len(where_crm) > 0 and len(where_company) > 0:
            where_clause = f'where {where_company} and {where_crm} '
        elif len(where_crm) > 0:
            where_clause = f'where {where_crm} '
        elif len(where_company) > 0:
            where_clause = f'where {where_company} '

        request = f"""
                    select 
                        sale.id,
                        sale.name,
                        sale.origin,
                        sale.opportunity_id as vkpid,
                        sale.state,
                        sale.user_id,
                        (case
                            when partner."name" > '' and partner.company_name > '' then
                                case 
                                    when partner."name" like partner.company_name then partner."name"
                                    else concat(partner."name",', ',partner.company_name)
                                end
                            else partner."name"
                        end) as partner,
                        sale.partner_id,
                        sale.machine_id ,
                        ms.model as typeid,
                        ms.mx_number as mx,
                        (replace(replace(replace((concat('[',trim(both ms.mx_number),']_',trim(both hst."name"),'_',trim(type."name"))),'#','_'),'/','_'),'\','_')) as ms_name,
                        type.name as type,
                        ms.manufacturer as hstid,
                        hst."name" as hst,
                        sale.company_id as firmaid,
                        firma."name" as firma,
                        ms.categ_id as msartid,
                        cat."name"  as msart,
                        ms.sub_categ_id as mssubartid,
                        subcat."name" as mssubart
                    from sale_order as sale
                        left join res_partner as partner on sale.partner_id = partner.id  
                        left join machine_list_machine as ms on sale.machine_id = ms.id 
                        left join machine_list_model as type on ms.model = type.id
                        left join machine_list_manufacturer as hst on ms.manufacturer = hst.id 
                        left join machine_list_category as cat on ms.categ_id = cat.id 
                        left join machine_list_sub_category as subcat on ms.sub_categ_id =subcat.id 
                        left join res_company as firma on sale.company_id = firma.id 
                    {where_clause} 
                    order by sale.id 	
                    """

        # print(request)
        # Limit und Offset
        if rowlimit != 0:
            request = request + f'\nlimit {rowlimit} offset {offset}'
        request = request + ";"
        # print(request)
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            return True, 'Sales-Liste ausgelesen', rows
        except:
            return False, (
                'Die Sales-Liste konnte nicht erfogreich abgefragt werden!'), None

    def get_sales_attachment_folder_link(self, sale_id):
        """
        Ermittelt, ob ein Eintrag '__(SPO)__Ordner' für das Sale
        mit der id sale_id giebt.
            sale_id   = Odoo_Id des Sale

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Liste mit den 
                      Eintragsdaten oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (
                    f'Es konnte keine Verbindung zur Datenbank {self._db}' +
                    f'auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                select * from ir_attachment 
                    where 
                        name = '__(SPO)_Ordner' and
                        res_model = 'sale.order' and
                        res_id = {sale_id};
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows
            else:
                return False, (
                        f'Keinen Ordner-Attachment-Eintrag für das Sale' +
                        f' {sale_id} gefunden!')
        except:
            return False, (
                'Die Attachmentliste konnte nicht erfogreich abgefragt werden!')

    def get_sales_attachment_file_links(
            self, sale_id, attachment_type=None):
        """
        Ermittelt die Dateianhänge für ein sale.
        (Die Eintrage für die Ordnerlinks werden ausgelassen!)
            sale_id         = Odoo_Id der Maschine
            attachment_type = Type des Anhangs
                                binary: local auf dem Odoo-Server gespeichert
                                url:    als Link eingetragen
                                keine Angabe: alle Einträge

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Liste der 
                      Einträge oder Fehlermeldung.
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                select * from ir_attachment 
                    where 
                        name != '__(SPO)_Ordner' and
                        res_model = 'sale.order' and
                        res_id = {sale_id}
                """
        if attachment_type != None:
            request = request + f" and type = '{attachment_type}'"

        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows
            else:
                return False, (f'Keinen Attachment-Einträge für das Sale ' +
                               f'{sale_id} gefunden!')
        except:
            return False, (
                'Die Attachmentliste konnte nicht erfogreich abgefragt werden!')

    def create_sales_attachment_folder_link(
            self, sale_id, sale_name, link_url, firma_id, user_id=1):
        """
        Erstellt einen Ordner-Attachment-Link für ein CRM.
            sale_id         = Odoo_Id des CRM
            sale_name       = Name des CRM in Form 
                              '[<mxnr>] <hersteller> <type>'
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                insert into ir_attachment (
                    name, 
                    datas_fname,
                    res_name,
                    res_model,
                    res_model_name,
                    res_id,
                    company_id,
                    type,
                    url,
                    mimetype,
                    active,
                    create_uid,
                    create_date,
                    write_uid,
                    write_date)
                values (
                    '__(SPO)_Ordner',
                    '__(SPO)_Ordner',
                    '{sale_name}',
                    'sale.order',
                    'Sale Order',
                    {sale_id},
                    {firma_id},
                    'url',
                    '{link_url}',
                    'application/spo.folder',
                    true,
                    {user_id},
                    now(),
                    {user_id},
                    now())
                returning *;
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                print(rows[0])
                return True, rows[0]['id']
            else:
                return False, (f'Konte keinen Folder-Attachment-Eintrag für ' +
                               f'den Sale {sale_id} erstellen!')
        except:
            return False, (f'Konte keinen Folder-Attachment-Eintrag für den ' +
                           f'Sale {sale_id} erstellen!')

    def create_sales_attachment_file_link(
            self, sale_id, sale_name, file_name, link_url, firma_id, mimetype, user_id=1):
        """
        Erstellt einen Ordner-Attachment-Link für eine Maschine.
            sale_id         = Odoo_Id des Sale
            sale_name       = Names des Sale im Odoo
            file_name       = Name der Datei im SPO.
                              Im Link wird dem Namen noch "_(SPO)_" 
                              vorangestellt. Dies kennzeichnet,
                              dass sich der Anhang im Sharepoint befindet.
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            mimetype        = Mimetype für den Anhang z.B. "application/pdf" 
                              o. "image/png"
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass,
                port=self._port, cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                insert into ir_attachment (
                    name, 
                    datas_fname,
                    res_name,
                    res_model,
                    res_model_name,
                    res_id,
                    company_id,
                    type,
                    url,
                    mimetype,
                    active,
                    create_uid,
                    create_date,
                    write_uid,
                    write_date)
                values (
                    '_(SPO)_{file_name}',
                    '{file_name}',
                    '{sale_name}',
                    'sale.order',
                    'Sale Order',
                    {sale_id},
                    {firma_id},
                    'url',
                    '{link_url}',
                    '{mimetype}',
                    true,
                    {user_id},
                    now(),
                    {user_id},
                    now())
                returning id;
                """
        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte keinen File-Attachment-Eintrag für ' +
                               f'den Sale {sale_id} erstellen!')
        except:
            return False, (f'Konte keinen File-Attachment-Eintrag für ' +
                           f'den Sale {sale_id} erstellen!')

    def update_sales_attachment_folder_link(
            self, sale_id, sales_name, link_url,
            firma_id, user_id=1, attachment_id=None):
        """
        Ändert einen Ordner-Attachment-Link für ein Sale.
        Die Filterung erfolgt entweder über die res_id="sale_id" und 
        res_model oder über die id="attachment_id".
        !! Achtung !! ist die "attachment_id" angegeben wird nur 
        über sie gefiltert !!
        Es werden die Felder res_name, company_id und url gesetzt. 
            sale_id         = Odoo_Id des Sale
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird 
                              die ID nicht angegeben, wird 1 (System) benutzt!
            attachment_id (opt.)    = Odoo-ID des Anhangeintrags

            Rückgabe: Liste, erster Wert True/False, Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                update ir_attachment 
                set
                    res_name = '{sale_name}',
                    company_id  = '{firma_id}',
                    url = '{link_url}',
                    active = true,
                    write_uid = '{user_id}',
                    write_date = now()
                where
                    res_model = 'sale.order' and
                    res_id = {sale_id} and
                    name = '__(SPO)_Ordner'
                returning id;
                """
        if attachment_id != None:
            request = f"""
                    update ir_attachment 
                    set
                        res_name = '{sales_name}',
                        company_id  = '{firma_id}',
                        res_id = {sale_id},
                        url = '{link_url}',
                        active = true,
                        write_uid = '{user_id}',
                        write_date = now()
                    where
                        id = {attachment_id}
                    returning *;
                    """

        # Abfrage ausführen
        cursor = conn.cursor()
        try:
            cursor.execute(request)
            conn.commit()
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            # Bei Bedarf, SSH-Tunnel schließen
            if self.__use_ssh_tunnel == True:
                self.__ssh_tunnel.stop()
            if len(rows) > 0:
                return True, rows[0]['id']
            else:
                return False, (f'Konte den Folder-Attachment-Eintrag für ' +
                               f'das Sale {sale_id} nicht ändern!')
        except:
            return False, (f'Konte den Folder-Attachment-Eintrag für das ' +
                           f'Sale {sale_id} nicht ändern!')

    def update_sales_attachment_file_link(
            self, attachment_id, sale_id, file_name, link_url,
            firma_id, mimetype, user_id=1):
        """
        Aktualisiert einen Ordner-Attachment-Link für eine Maschine.
        !! Achtung auch Anhangeinträge vom Type "binary" werden in einen 
        Link umgewandelt !!
            attachment_id   = Odoo-Id des Anhangeintrags
                              Die Filterung erfolgt nur über die ID!! 
            sale_id         = Odoo_Id des Sale
            file_name       = Name der Datei im SPO. Im Link wird dem Namen 
                              nach "_(SPO)_" vorangestellt. Dies kennzeichnet,
                              dass sich der Anhang im Sharepoint befindet.
            link_url        = Die URL des Docset's der Maschine.
            firma_id        = ID der Firma im Odoo
            mimetype        = Mimetype für den Anhang z.B. "application/pdf" 
                              o. "image/png"
            user_id (opt.)  = User-Id für Schreib- und Änderungsbenutzer. Wird
                              die ID nicht angegeben, wird 1 (System) benutzt!

            Rückgabe: Liste, erster Wert True/False, zweiter Wert die ID des 
                      Eintrags oder Fehlermeldung
        """

        try:
            # Bei Bedarf, SSH-Tunnel zur DB aufbauen
            if self.__use_ssh_tunnel == True:
                try:
                    self.__ssh_tunnel.start()
                except:
                    return False, ('Kann den SSH-Tunnel für die ' +
                                   'Datenbanverbindung nicht aufbauen!')
            # datenbankverbindung aufbauen
            conn = psycopg2.connect(
                host=self._host, dbname=self._db, user=self._user,
                password=self.__pass, port=self._port,
                cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            return False, (f'Es konnte keine Verbindung zur Datenbank ' +
                           f'{self._db} auf {self._host} aufgebaut werden!')

        # Select-Abfrage für die Liste
        request = f"""
                    update ir_attachment 
                    set
                        name = '_(SPO)_{file_name}',
                        datas_fname = '{file_name}',
                        company_id  = '{firma_id}',
                        res_id = {sale_id},
                        url = '{link_url}',
                        type = 'url',
                        active = true,
                        write_uid = '{user_id}',
                        write_date = now(),
                        mimetype='{mimetype}'
                    where
                        id = {attachment_id}
                    returning *;
                    """
        # print(request)
        # Abfrage ausführen
        cursor = conn.cursor()
        # try:
        cursor.execute(request)
        conn.commit()
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        # Bei Bedarf, SSH-Tunnel schließen
        if self.__use_ssh_tunnel == True:
            self.__ssh_tunnel.stop()
        if len(rows) > 0:
            # print(rows[0])
            return True, rows[0]['id']
        else:
            return False, (f'Konte den File-Attachment-Eintrag für das ' +
                           f'Sale {sale_id} nicht ändern!')
