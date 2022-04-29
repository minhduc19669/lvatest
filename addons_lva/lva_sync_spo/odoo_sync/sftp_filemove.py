""" 
Dise Datei ist Bestandteil der Dateisyncronisation
zwischen odoo und SPO. Dabei geht es darum, zu einem
entsprechenden Objekt im Odoo ein entsprechendes Docsets
anzulegen und vorhandene dateien dort abzulegen.
Sie stellt die Objekte und Funktionen bereit, 
die Dateien per sftp verschieben.
"""

from paramiko import SSHClient, AutoAddPolicy, ssh_exception
import os


class SSHsftp:
    """ 
    Diese Klasse diehnt den Zugriff auf den Filstore von Odoo per ssh/sftp 
    """

    def __init__(self, ssh_host, ssh_user, ssh_pass, ssh_port=22):
        # Zugangsdaten zum SSH-Server
        self._host = ssh_host
        self._user = ssh_user
        self.__pass = ssh_pass
        self._port = ssh_port
        # Pfadangaben für den Dateitranfer
        self._remote_base_path = ''
        self._local_temp_path = ''

    def test_conn(self):
        # Testet den SSH-Zugang

        try:
            ssh = SSHClient()
            ssh.set_missing_host_key_policy(AutoAddPolicy())
            ssh.connect(
                hostname=self._host, username=self._user,
                password=self.__pass, port=self._port)
            ssh.close()
            return True, (f'Die Verbindung zum SSH-Server {self._host} ' +
                          'konnte aufgebaut werden.')
        except ssh_exception.AuthenticationException:
            return False, (f'Die Authentifikation des Benutzers {self._user} ' +
                           f'am SSH-Server {self._host} ist fehlgeschlagen!')
        except:
            return False, (f'Es konnte keine Verbindung mit dem SSH-Server ' +
                           f'{self._host} aufgebaut werden')

    def set_base_path(self, local_path='', remote_path=''):
        """
        Setzt die Basispfade für den lokalen bzw. 
        entfernten Dateizugriff. 
            local_path (opt.)   = lokaler Basispfad
            remote_path (opt.)  = entfernter Basipfad
        """

        if local_path != '':
            self._local_temp_path = local_path

        if remote_path != '':
            self._remote_base_path = remote_path

    def copy_odoo_to_temp(self, odoo_rel_path, local_filename=''):
        """
        Kopiert eine Datei aus dem Odoo-Filestore
        in das locale Temp-Verzeichniss
            odoo_rel_path   = Dateipfad relativ zum entfernten Basisordner 
                              (Angabe, wie sie in der Datenbank steht)
            local_filename  = Dateiname für die lokale Ablage

            Rückgabe: Liste erster Wert True/False, Zweiter Wert Fehlermeldung
        """
        ##try:
        try:
            ssh = SSHClient()
            ssh.set_missing_host_key_policy(AutoAddPolicy())
            ssh.connect(
                hostname=self._host, username=self._user,
                password=self.__pass, port=self._port)
            print(self._host, self._user, self.__pass, self._port)
            sftp = ssh.open_sftp()
            remote_file = (
                f"/{self._remote_base_path}/{odoo_rel_path}").replace('//', '/')
            if local_filename == '':
                local_filename = os.path.basename(odoo_rel_path)
            if self._local_temp_path != '':
                local_file = (
                    f'/{self._local_temp_path}/{local_filename}').replace('//', '/')
            else:
                local_file = local_filename
            #    try:
            print('--> copy ', remote_file, ' - ', local_file)
            sftp.get(remote_file, local_file)
            sftp.close()
            ssh.close()
        except:
            return False, 'not work'
        return True, f'Datei {remote_file} nach {local_file} kopiert.'
        #    except:
        #        ssh.close()
        #        return False, (f'Konnte Datei {remote_file} nicht ' +
        #                       f'nach {local_file} kopieren!')

        # except paramiko.ssh_exception.AuthenticationException:
        #    return False, (
        #        f'Die Authentifikation des Benutzers {self._user} ' +
        #        f'am SSH-Server {self._host} ist fehlgeschlagen!')
        # except:
        #    return False, (f'Es konnte keine Verbindung mit dem SSH-Server ' +
        #                   f'{self._host} aufgebaut werden')

    def move_odoo_to_temp(self, odoo_rel_path, local_filename=''):
        """
        Verschiebt eine Datei aus dem Odoo-Filestore
        in das locale Temp-Verzeichniss
            odoo_rel_path   = Dateipfad relativ zum entfernten Basisordner 
                              (Angabe, wie sie in der Datenbank steht)
            local_filename  = Dateiname für die lokale Ablage

            Rückgabe: Liste erster Wert True/False, Zweiter Wert Fehlermeldung
        """
        try:
            ssh = SSHClient()
            ssh.set_missing_host_key_policy(AutoAddPolicy())
            ssh.connect(
                hostname=self._host, username=self._user,
                password=self.__pass, port=self._port)
            sftp = ssh.open_sftp()
            remote_file = (
                f"/{self._remote_base_path}/{odoo_rel_path}").replace('//', '/')
            if local_filename == '':
                local_filename = os.path.basename(odoo_rel_path)
            if self._local_temp_path != '':
                local_file = (
                    f'/{self._local_temp_path}/{local_filename}').replace(
                    '//', '/')
            else:
                local_file = local_filename
            try:
                # Datei in den Temp-Ordner kopieren
                sftp.get(remote_file, local_file)
                sftp.close()
                try:
                    # Datei vom Remote-Ordner löschen
                    stdin, stdout, stderr = ssh.exec_command(
                        f'rm -f {remote_file}')
                    error_lines = stderr.readlines()
                    error_lines_count = len(error_lines)
                    error_text = "".join(error_lines)
                    if error_text[-1] == '\n':
                        error_text = error_text[:-1]
                    out_lines = stdout.readlines()
                    out_lines_count = len(out_lines)
                    out_text = "".join(out_lines)
                    ssh.close()
                    if out_text[-1] == '\n':
                        out_text = out_text[:-1]
                    # Bei Fehler bei der Befehlsausführung
                    if error_lines_count != 0:
                        return False, (
                            f'Datei {remote_file} wurde nach {local_file} ' +
                            f'kopiert aber nicht gelöscht!', error_text)
                    else:
                        return True, (f'Datei {remote_file} wurde erfogreich ' +
                                      'nach {local_file} verschoben!')
                except:
                    ssh.close()
                    return False, (f'Datei {remote_file} wurde nach ' +
                                   f'{local_file} kopiert aber nicht gelöscht!')
            except:
                ssh.close()
                return False, (f'Konnte Datei {remote_file} nict nach ' +
                               f'{local_file} kopieren!')
        except ssh_exception.AuthenticationException:
            return False, (
                    f'Die Authentifikation des Benutzers {self._user} am ' +
                    f'SSH-Server {self._host} ist fehlgeschlagen!')
        except:
            return False, (f'Es konnte keine Verbindung mit dem SSH-Server ' +
                           f'{self._host} aufgebaut werden')
