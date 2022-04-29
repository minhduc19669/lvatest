""" 
Dise Datei ist Bestandteil der Dateisyncronisation
zwischen odoo und SPO. Dabei geht es darum, zu einem
entsprechenden Objekt im Odoo ein entsprechendes Docsets
anzulegen und vorhandene dateien dort abzulegen.
Sie stellt die Objekte und Funktionen bereit, 
die für die Verbindung mit dem SPO benötigt werden.
"""

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.caml.caml_query import CamlQuery
from office365.sharepoint.documentmanagement.document_set import DocumentSet
from office365.sharepoint.contenttypes.content_type import ContentType
import os
from urllib.parse import urlparse


# from office365.sharepoint.listitems.listitem.ListItem import f

class SPOLibrary:
    """ Diese Klasse diehnt der Handhabung einer Dokumentbibliothek im SPO """

    def __init__(self, website, library, user, passwd):
        """ 
        Hinterlegt die Grunddaten für den Zugriff auf eine SPO-Bibliothek

            website  - Die Website, in der sich die Bibliothek befindet
            library  - Der Titel der Bibliothek
            user     - Der Benutzer für den Zugriff 
            passwd   - Das password für den Zugriff
        """
        self._website = website
        self._library = library
        self._user = user
        self.__passwd = passwd

    def test_connection(self):
        """
        Testet, ob die Verbindung aufgebaut werden kann.

        Rückgabe: Liste, erster Eintrag True o. Fals, zweiter Eintrag Fehlertext
        """
        error = None
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            error = 'False, Keine verbindung zur Website!'
            return error

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            error = "False, Keine verbindung zur Bibliothek!"
            return error

        # Zugriff erfolgreich
        error = "Connection Success !"
        return error

    def get_folder_id_by_path(self, folder_path):
        """
        Giebt die ID des Folders/Docsets zurück, dessen Pfad angegeben wurde.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            folder_path = Pfadangabe des Ordners innerhalb der Bibliothek
                          z.B. /hauptordner/suchordner(bei Unterverzeichniss), 
                          suchordner (Ordner auf Root-Ebene)

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='Title'/>
                                            <Value Type='Text'>{os.path.basename(folder_path)}</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                                <ViewFields>
                                    <FieldRef Name='Title'/>
                                    <FieldRef Name='ID'/>
                                </ViewFields>
                            </View>"""

        # Bei Unterordner, die ServerRelativUrl angeben
        if os.path.dirname(folder_path).replace('/', '') != '':
            spo_query.FolderServerRelativeUrl = (
                    urlparse(self._website).path + '/' + self._library + '/'
                    + os.path.dirname(folder_path)).replace('//', '/')

        # Request
        spo_items = list.get_items(spo_query).execute_query()
        if len(spo_items) > 0:
            return True, spo_items[0]._properties['ID']
        else:
            return False, f'Konnte Folder/Docset "{folder_path}" nicht finden'

    def get_folder_by_path(self, folder_path):
        """
        Giebt die Daten des Folders/Docsets zurück, dessen Pfad angegeben wurde.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            folder_path = Pfadangabe des Ordners innerhalb der Bibliothek
                          z.B. /hauptordner/suchordner(bei Unterverzeichniss), 
                          suchordner (Ordner auf Root-Ebene)

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die Folderdaten
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='FileLeafRef'/>
                                            <Value Type='Text'>{os.path.basename(folder_path)}</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                                <ViewFields>
                                    <FieldRef Name='Title'/>
                                    <FieldRef Name='ID'/>
                                </ViewFields>
                            </View>"""
        # Bei Unterordner, die ServerRelativUrl angeben
        if os.path.dirname(folder_path).replace('/', '') != '':
            spo_query.FolderServerRelativeUrl = (urlparse(
                self._website).path + '/' + self._library
                                                 + '/' + os.path.dirname(folder_path)).replace('//', '/')
        spo_items = list.get_items(spo_query).execute_query()

        # path = "/home/minhduc/LVA/sync_odoo_spo/odoo.conf"
        # with open(path, 'rb') as content_file:
        #     file_content = content_file.read()
        # list_title = "dsdsd"
        # site_url = "https://bnksolution.sharepoint.com/sites/ducluong"
        # ctx = ClientContext(site_url).with_credentials(UserCredential(self._user, self.__passwd))
        # # web = ctx.web
        # # ctx.load(web)
        # # ctx.execute_query()
        # # print("Web title: {0}".format(web.properties['Title']))
        # target_folder = ctx.web.lists.get_by_title(list_title).root_folder
        # name = os.path.basename(path)
        # target_file = target_folder.upload_file(name, file_content).execute_query()
        # print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))

        if len(spo_items) > 0:
            return True, spo_items[0]
        else:
            return False, f'Konnte Folder/Docset "{folder_path}" nicht finden'

    def get_maschdocsets(
            self, contenttype, id=None, hersteller_id=None, maschart_id=None
            , firma_id=None, odoo_id=None, bestand=None, name=None, title=None):
        """
        Liest die vorhandenen Maschinendocsets (ContentType="Maschinenmappe") 
        aus und giebt diese als Liste zurück. Die Liste kann mittels der 
        optionalen Parameter gefiltert werden. Mehrere Filter werden mit 
        UND verknüpft!
            id (opt.)               = ID in der Bibliothek
            hersteller_id (opt.)    = Hersteller-Id im Odoo
            maschart_id (opt.)      = Maschinenart-id (Category) im Odoo
            firma_id (opt.)         = ID Der Firma im Odoo
            odoo_id (opt.)          = ID der Maschine im Odoo
            bestand                 = boolscher Wert, ob Maschine im Bestand
            name                    = Name(nicht der Titel) der Documentenmappe 
            title                   = Titel der Documentenmappe

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder Liste der gefundenen Docsets
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Filter erstellen
        where = f"""
                    <Eq>
                        <FieldRef Name='ContentType' />
                        <Value Type='Text'>{contenttype}</Value>
                    </Eq>
              """
        if id != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="ID" />
                            <Value Type='Text'>{id}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if hersteller_id != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="Hersteller_id" />
                            <Value Type='Integer'>{hersteller_id}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if maschart_id != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="Maschinenart_id" />
                            <Value Type='Integer'>{maschart_id}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if firma_id != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="Firmen_id" />
                            <Value Type='Integer'>{firma_id}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if odoo_id != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="odoo_id" />
                            <Value Type='Integer'>{odoo_id}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if bestand != None:
            bool_best = 0
            if bestand == True:
                bool_best = 1
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="Bestand" />
                            <Value Type='Boolean'>{bool_best}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if name != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="FileLeafRef" />
                            <Value Type='File'>{name}</Value>
                        </Eq>
                        {where}
                    </And>
                   """
        if title != None:
            where = f"""
                    <And>
                        <Eq>
                            <FieldRef Name="Title" />
                            <Value Type='Text'>{title}</Value>
                        </Eq>
                        {where}
                    </And>
                   """

        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        {where}
                                    </Where>
                                </Query>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) > 0:
            return True, spo_items
        else:
            return False, f'Konnte kein/e Docset/s finden'

    def get_maschdocset_id_by_odoomaschid(self, odoo_masch_id):
        """
        Giebt die ID des Docsets einer Maschinenbibliothek zurück, 
        dessen Feld 'odoo_id' der odoo_masch_id entspricht.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            odoo_masch_id = Die Odoo-Id der Maschine, zu der das Docset gehört

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='ContentType' />
                                                <Value Type='Text'>Maschinenmappe</Value>
                                            </Eq>
                                            <Eq>
                                                <FieldRef Name='odoo_id'/>
                                                <Value Type='Integer'>
                                                    {odoo_masch_id}
                                                </Value>
                                            </Eq>
                                        </And>
                                    </Where>
                                </Query>
                                <ViewFields>
                                    <FieldRef Name='Title'/>
                                    <FieldRef Name='ID'/>
                                </ViewFields>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) > 0:
            return True, spo_items[0]._properties['ID']
        else:
            return False, (
                f'Konnte kein Docset mit der odoo_id "{odoo_masch_id}" finden')

    def get_maschdocset_by_odoomaschid(self, odoo_masch_id):
        """
        Giebt die Daten des Docsets einer Maschinenbibliothek zurück, 
        dessen Feld 'odoo_id' der odoo_masch_id entspricht.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            odoo_masch_id = Die Odoo-Id der Maschine, zu der das Docset gehört

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die Docset-Daten
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='ContentType' />
                                                <Value Type='Text'>
                                                    Maschinenmappe
                                                </Value>
                                            </Eq>
                                            <Eq>
                                                <FieldRef Name='odoo_id'/>
                                                <Value Type='Integer'>
                                                    {odoo_masch_id}
                                                </Value>
                                            </Eq>
                                        </And>
                                    </Where>
                                </Query>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) > 0:
            return True, spo_items[0]
        else:
            return False, (
                f'Konnte kein Docset mit der odoo_id "{odoo_masch_id}" finden')

    def get_crmdocset_id_by_odoocrmid(self, odoo_crm_id):
        """
        Giebt die ID des Docsets einer CRM-bibliothek zurück, 
        dessen Feld 'CRM_id' der odoo_crm_id entspricht.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            odoo_crm_id = Die Odoo-Id des CRM, zu der das Docset gehört

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        querystring = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='ContentType' />
                                                <Value Type='Text'>VK_Projekt</Value>
                                            </Eq>
                                            <Eq>
                                                <FieldRef Name='CRM_id'/>
                                                <Value Type='Integer'>
                                                    {odoo_crm_id}
                                                </Value>
                                            </Eq>
                                        </And>
                                    </Where>
                                </Query>
                                <ViewFields>
                                    <FieldRef Name='Title'/>
                                    <FieldRef Name='ID'/>
                                </ViewFields>
                            </View>"""
        spo_query.ViewXml = querystring
        # print('oooo>>>  ',querystring)
        # Request
        spo_items = list.get_items(spo_query).execute_query()
        if len(spo_items) > 0:
            return True, spo_items[0]._properties['ID']
        else:
            return False, (
                f'Konnte kein Docset mit der odoo_id "{odoo_crm_id}" finden')

    def get_crmdocset_by_odoocrmid(self, odoo_crm_id):
        """
        Giebt die Daten des Docsets einer CRM-Bibliothek zurück, 
        dessen Feld 'CRM_id' der odoo_crm_id entspricht.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            odoo_masch_id = Die Odoo-Id des CRM, zu der das Docset gehört

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die Docset-Daten
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        querystring = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='ContentType' />
                                                <Value Type='Text'>VK_Projekt</Value>
                                            </Eq>
                                            <Eq>
                                                <FieldRef Name='CRM_id'/>
                                                <Value Type='Integer'>
                                                    {odoo_crm_id}
                                                </Value>
                                            </Eq>
                                        </And>
                                    </Where>
                                </Query>
                            </View>"""
        spo_query.ViewXml = querystring
        # print('oooo>>>  ',querystring)
        # Request
        spo_items = list.get_items(spo_query).execute_query()
        # for item in spo_items:
        # print("--->>", item._properties)

        if len(spo_items) > 0:
            return True, spo_items[0]
        else:
            return False, (
                f'Konnte kein Docset mit der odoo_id "{odoo_crm_id}" finden')

    def get_salesdocset_id_by_odoosalesid(self, odoo_sales_id):
        """
        Giebt die ID des Docsets einer CRM-Bibliothek zurück, 
        dessen Feld 'Sales_id' der odoo_sales_id entspricht.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            odoo_crm_id = Die Odoo-Id des CRM, zu der das Docset gehört

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='ContentType' />
                                                <Value Type='Text'>VK_Angebotsmappe</Value>
                                            </Eq>
                                            <Eq>
                                                <FieldRef Name='Sales_id'/>
                                                <Value Type='Integer'>
                                                    {odoo_sales_id}
                                                </Value>
                                            </Eq>
                                        </And>
                                    </Where>
                                </Query>
                                <ViewFields>
                                    <FieldRef Name='Title'/>
                                    <FieldRef Name='ID'/>
                                </ViewFields>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) > 0:
            return True, spo_items[0]._properties['ID']
        else:
            return False, (
                f'Konnte kein Docset mit der odoo_id "{odoo_sales_id}" finden')

    def get_salesdocset_by_odoosalesid(self, odoo_sales_id):
        """
        Giebt die Daten des Docsets einer CRM-Bibliothek zurück, 
        dessen Feld 'Sales_id' der odoo_sales_id entspricht.
        Wird der Ordner nicht gefunden, wird False zurückgegeben.

            odoo_sales_id = Die Odoo-Id des Sale, zu der das Docset gehört

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die Docset-Daten
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> ID des Folders/Docsets ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='ContentType' />
                                                <Value Type='Text'>VK_Angebotsmappe</Value>
                                            </Eq>
                                            <Eq>
                                                <FieldRef Name='Sales_id'/>
                                                <Value Type='Integer'>
                                                    {odoo_sales_id}
                                                </Value>
                                            </Eq>
                                        </And>
                                    </Where>
                                </Query>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) > 0:
            return True, spo_items[0]
        else:
            return False, (
                f'Konnte kein Docset mit der odoo_id "{odoo_sales_id}" finden')

    def get_item_by_id(self, id, select=None):
        """
        Giebt die Daten eines Items an Hand desen ID zurück, 
        Wird das Item nicht gefunden, wird False zurückgegeben.

            id              = Die ID des Items
            select (opt.)   = Liste der auszugebenen Felder

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die Item-Daten
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> Item ermitteln <<====
        spo_items = list.get_items()
        # for property, value in vars(spo_items._query_options).items():
        #    print(property, ":", value)

        if select != None:
            spo_items._query_options.select = select
        spo_items._query_options.filter = f"ID eq {id}"
        spo_ctx.load(spo_items)
        spo_ctx.execute_query()
        # print("****   ",len(spo_items))
        if len(spo_items) > 0:
            return True, spo_items[0]
        else:
            return False, f'Konnte kein Item mit der ID "{id}" finden'

    def set_data_to_item(self, item_id, field_data):
        """
        Setzt die Metadaten zu einem sharepoint-Item.

            item_id     = Die ID des Eintrags/Folders/Docsets/Files 
                          in der Bibliothek/Liste
            field_data  = Ein Dictionary mit den zu setzenden SPO-Feldern und 
                          deren Daten. Nicht angegebene Felder werden auch 
                          nicht geändert.

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """
        # print(field_data)
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> Eintrags/Folders/Docsets/Files anhand der ID ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='ID'/>
                                            <Value Type='Text'>{item_id}</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) == 0:
            return False, (
                    f'Konnte kein Eintrags/Folders/Docsets/Files mit der ID ' +
                    f'"{item_id}" finden!')
        else:
            error_filds = []
            for field in field_data:
                try:
                    # print(field, field_data[field])
                    spo_items[0].set_property(
                        field, field_data[field]).update().execute_query()
                except:
                    # Fehlerhafte Felder für Fehlermeldung in Liste eintragen
                    error_filds.append(field)
            if len(error_filds) == 0:
                return True, (
                        f'Eintrag/Folder/Docset/File mit der ID "' +
                        f'{item_id}" erfolgreich geändert!')
            else:
                error_text = '; '.join(error_filds)
                return True, ((
                                      f'Eintrag/Folder/Docset/File mit der ID "{item_id}" ' +
                                      f'geändert! Aber Felder {error_text} ausgelassen!')
                , error_filds)

    def create_docset(self, folder_path, field_data=None):
        """
        Erstellt ein Docset in der Bibliothek.

            folder_path         =   Pfadangabe des Docsets innerhalb der 
                                    Bibliothek z.B. /hauptordner/docset(bei 
                                    Unterverzeichniss), docset (Ordner auf 
                                    Root-Ebene)
            field_data (opt.)   =   Ein Dictionary mit den zu setzenden 
                                    SPO-Feldern und deren Daten. Nicht 
                                    angegebene Felder werden auch nicht 
                                    geändert.

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """

        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> Docset erstellen <<====

        # Wenn vorhanden, führenden '/' aus folder_path entfernen
        if folder_path[0] == '/':
            folder_path = folder_path[1:]

        # Docset erstellen
        try:
            docset = DocumentSet.create(
                spo_ctx, list.root_folder, folder_path).execute_query()
            # for property, value in vars(docset).items():
            #    print(property, ":", value)

        except:
            return False, f'Docset "{folder_path}" konnte nicht angelegt werden'

        # print(f"docset id : {docset.properties['ID']}")
        if 'ID' in docset.properties:
            # wenn Docset erstellt werden konnten der
            # Bei Bedarf Metadaten für Docset setzen
            if field_data != None:
                result = self.set_data_to_item(
                    docset.properties['ID'], field_data)
                if result[0] == True:
                    return True, docset, (
                            f'Docset wurde mit ID {docset.properties["ID"]} ' +
                            f'angelegt. {result[1]} und die Metadaten eingetragen.')
                else:
                    return True, docset, (
                            f'Am Docset "{folder_path}" konnte die Felder ' +
                            f'nicht gesetz werden!\n{result[1]}')
            else:
                return True, docset, (
                        f'Docset wurde mit ID {docset.properties["ID"]} ' +
                        f'angelegt. Es waren keine Metadaten angegeben.')
        else:
            return False, (
                f'Fehler! Das Docset "{folder_path}" konnte '
                f'nicht angelegt werden!')

    def create_folder(self, folder_path):
        """
        Erstellt ein Folder in der Bibliothek. Wenn nicht vorhanden, werden 
        auch die übergeordneten Folder mit erstellt.

            folder_path         =   Pfadangabe des Docsets innerhalb der 
                                    Bibliothek z.B. /hauptordner/docset(bei 
                                    Unterverzeichniss), docset 
                                    (Ordner auf Root-Ebene)

            Rückgabe: Liste, erster Wert True/False, zweiter Wert 
                      Fehlermeldung oder die ID
        """

        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # Folder erstellen
        full_folder_path = f"/{self._library}/{folder_path}".strip(
        ).replace('//', '/')
        folder = spo_ctx.web.ensure_folder_path(full_folder_path).execute_query()
        # Folder-ID ermitteln
        result = self.get_folder_id_by_path(
            f"/{folder_path}".strip().replace('//', '/'))
        if result[0] == False:
            return False, (
                "Ordner konnte nicht sicher erstellt werden!{result[1]}")

        return True, result[1]

    def upload_file_to_library(
            self, local_path, target_path=None, field_data=None):
        """
        Kopiert eine Datei in eine Sharepoint-Bibliothek
            local_path          = der vollständige Pfad der lokalen Datei 
                                    z.B. /tmp/tempfile.txt
            target_path (opt.)  = Zielpfad der Datei in der Bibliothek. Wird er
                                  nicht angegeben, wird die Datei mit dem dem
                                  ursprünglichen Dateinamen ins Root-Verzeichnis
                                  der Bibliothek kopiert. Wenn nur Zielordner 
                                  angegeben werden soll, so ist die Angabe mit 
                                  "/" abzuschließen.
                                  Beispiele:
                                  /folder/          = Datei wird mit 
                                                      urspünglichen Namen ins 
                                                      Verzeichnis/Docset 
                                                      "folder" kopiert
                                  /folder2/test.txt = Datei wird mit Namen 
                                                      "test.txt" ins ins 
                                                      Verzeichnis/Docset 
                                                      "folder2" kopiert
            field_data  (opt.)  = Ein Dictionary mit den zu setzenden 
                                  SPO-Feldern und deren Daten. Nicht angegebene 
                                  Felder werden auch nicht geändert.
            
            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder URL der Datei
        """

        try:
            with open(local_path, 'rb') as file:
                content = file.read()
        except:
            return False, (
                f'Kann die lokale Datei "{local_path}" nicht einlesen.')

        # Dateinamen aus Quelle ermitteln
        filename = os.path.basename(local_path)

        # Pfad und Dateiangaben des Ziels prüfen
        target_folder = f'/{self._library}'
        if target_path != None:
            # feststellen, ob Zielpfad auch Dateinamen enthält
            if target_path[-1] != '/':
                # Dateinamen auf Zielangabe setzen
                filename = os.path.basename(target_path)
                # Zielverzeichnis ergänzen
                subfolder = os.path.dirname(target_path)
                target_folder = f'{target_folder}/{subfolder}'.replace(
                    '//', '/')
                if target_folder[-1] == '/':
                    target_folder = target_folder[:-1]

        # -> Dateiupload <-
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._library)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        spo_target = spo_ctx.web.ensure_folder_path(target_folder)
        try:
            spo_file = spo_target.upload_file(
                filename, content).execute_query()
        except:
            return False, (
                    f'Konnte Datei {filename} nicht nach {target_folder} ' +
                    f'hochladen!')

        # URL der Datei aus LinkingUri (Officedokumente)
        file_url = spo_file._properties['LinkingUri']
        # Wenn LinkingUri leer dan link aus ServerRelativUrl Bilden
        if file_url == None:
            url_data = urlparse(self._website)
            file_url = (f'{url_data.scheme}://{url_data.netloc}' +
                        '{spo_file._properties["ServerRelativeUrl"]}')

        # -> Metadaten setzen <-
        if field_data != None:
            error_fields = []
            spo_item = spo_file.listItemAllFields.execute_query()
            for field in field_data:
                try:
                    spo_item.set_property(
                        field, field_data[field]).execute_query()
                except:
                    error_fields.append(field)

            if len(error_fields) == 0:
                return True, file_url, (
                        f'Die Datei {filename} wurde erfolgreich nach ' +
                        f'{target_folder} hochgeladen und die Metadaten gesetz.')
            else:
                return True, file_url, (
                        f'Die Datei {filename} wurde erfolgreich nach ' +
                        f'{target_folder} hochgeladen.\n Die Metadaten ' +
                        f'{error_fields} konnten aber nicht gesetz werden!')
        else:
            return True, file_url, (
                    f'Die Datei {filename} wurde erfolgreich nach ' +
                    f'{target_folder} hochgeladen.')


class SPOList:
    """ Diese Klasse diehnt der Handhabung einer SPO-Liste """

    def __init__(self, website, list, user, passwd):
        """ 
        Hinterlegt die Grunddaten für den Zugriff auf eine SPO-Liste

            website  - Die Website, in der sich die Bibliothek befindet
            list     - Der Titel der Liste
            user     - Der Benutzer für den Zugriff 
            passwd   - Das password für den Zugriff
        """
        self._website = website
        self._list = list
        self._user = user
        self.__passwd = passwd

    def test_connection(self):
        """
        Testet, ob die Verbindung aufgebaut werden kann.

        Rückgabe: Liste, erster Eintrag True o. Fals, zweiter Eintrag Fehlertext
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._list)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Liste!"

        # Zugriff erfolgreich
        return True, "Verbindung erfogreich."

    def get_item_by_id(self, id, select=None):
        """
        Giebt die Daten eines Items an Hand desen ID zurück, 
        Wird das Item nicht gefunden, wird False zurückgegeben.

            id              = Die ID des Items
            select (opt.)   = Liste der auszugebenen Felder

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die Item-Daten
        """
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._list)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Bibliothek!"

        # ===>> Item ermitteln <<====
        spo_items = list.get_items()
        # for property, value in vars(spo_items._query_options).items():
        #    print(property, ":", value)

        if select != None:
            spo_items._query_options.select = select
        spo_items._query_options.filter = f"ID eq {id}"
        spo_ctx.load(spo_items)
        spo_ctx.execute_query()
        # print("****   ",len(spo_items))
        if len(spo_items) > 0:
            return True, spo_items[0]
        else:
            return False, f'Konnte kein Item mit der ID "{id}" finden'

    def set_data_to_item(self, item_id, field_data):
        """
        Setzt die Metadaten zu einem sharepoint-Item.

            item_id     = Die ID des Items 
                          in der Liste
            field_data  = Ein Dictionary mit den zu setzenden SPO-Feldern und 
                          deren Daten. Nicht angegebene Felder werden auch 
                          nicht geändert.

            Rückgabe: Liste, erster Wert True/False, zweiter Wert Fehlermeldung 
                      oder die ID
        """
        # print(field_data)
        spo_ctx = ClientContext(self._website).with_credentials(
            UserCredential(self._user, self.__passwd))
        # Zugriff auf Website testen
        try:
            web = spo_ctx.web.get().execute_query()
        except:
            return False, "Keine verbindung zur Website!"

        # Zugriff auf Bibliothek testen
        try:
            list = spo_ctx.web.lists.get_by_title(self._list)
            spo_ctx.load(list).execute_query()
        except:
            return False, "Keine verbindung zur Liste!"

        # ===>> Item anhand der ID ermitteln <<====
        # Caml Query erstellen
        spo_query = CamlQuery()
        spo_query.ViewXml = f"""
                            <View Scope='Recursive'>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='ID'/>
                                            <Value Type='Text'>{item_id}</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                            </View>"""

        # Request
        spo_items = list.get_items(spo_query).execute_query()

        if len(spo_items) == 0:
            return False, (
                    f'Konnte kein Item mit der ID ' +
                    f'"{item_id}" finden!')
        else:
            error_filds = []
            for field in field_data:
                try:
                    # print(field, field_data[field])
                    spo_items[0].set_property(
                        field, field_data[field]).update().execute_query()
                except:
                    # Fehlerhafte Felder für Fehlermeldung in Liste eintragen
                    error_filds.append(field)
            if len(error_filds) == 0:
                return True, (
                        f'Item mit der ID "' +
                        f'{item_id}" erfolgreich geändert!')
            else:
                error_text = '; '.join(error_filds)
                return True, ((
                                      f'Item mit der ID "{item_id}" ' +
                                      f'geändert! Aber Felder {error_text} ausgelassen!')
                , error_filds)

    def create_item(self, title, field_data=None):
        pass
