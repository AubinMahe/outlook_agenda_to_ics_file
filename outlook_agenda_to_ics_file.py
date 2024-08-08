#!/usr/bin/python3
#
import csv
from datetime import datetime
from io import TextIOWrapper
from typing import TextIO, Any
import sys
#
def create_header_corrected_CSV( agenda_outlook_csv_path: str ) -> TextIO:
   last_dot = agenda_outlook_csv_path.rfind('.')
   agenda_outlook_csv_header_corrected_path = f"{agenda_outlook_csv_path[0:last_dot]}-header-corrected.csv"
   with open( agenda_outlook_csv_path                 , "rt", encoding = 'WINDOWS-1252' ) as agenda_outlook_csv_file,\
        open( agenda_outlook_csv_header_corrected_path, "wt", encoding = 'UTF-8'        ) as agenda_outlook_csv_header_corrected_file:
      lines = agenda_outlook_csv_file.readlines()
      header = lines[0]
      header = header.replace( '"Début"', '"Début-le"', 1 )
      header = header.replace( '"Début"', '"Début-à"' , 1 )
      header = header.replace( '"Fin"'  , '"Fin-le"'  , 1 )
      header = header.replace( '"Fin"'  , '"Fin-à"'   , 1 )
      lines[0] = header
      agenda_outlook_csv_header_corrected_file.writelines( lines )
      agenda_outlook_csv_header_corrected_file.close()
      return open( agenda_outlook_csv_header_corrected_path, "rt" )
#
def fold( text: str ) -> str:
   global CRLF
   folded = text[0:75] + CRLF
   text   = text[75:]
   while len( text ) > 0:
      folded += '\t' + text[0:74] + CRLF
      text    = text[74:]
   return folded
#
def csv_to_ics( row: dict[str | Any, str | Any], ics_file: TextIOWrapper ) -> None:
   objet                     = row["Objet"]
   debut_date                = row["Début-le"]
   debut_hour                = row["Début-à"]
   fin_date                  = row["Fin-le"]
   fin_hour                  = row["Fin-à"]
   organisateur              = row["Organisateur d'une réunion"]
   participants_obligatoires = row["Participants obligatoires" ].replace(";",",")
   participants_facultatifs  = row["Participants facultatifs"  ].replace(";",",")
   emplacement               = row["Emplacement"]
   description               = row["Description"].strip().replace( '\n', '\\n' )
   debut = datetime.strptime( debut_date + '-' + debut_hour, "%d/%m/%Y-%H:%M:%S" ).strftime( "%Y%m%dT%H%M%S" )
   fin   = datetime.strptime( fin_date   + '-' + fin_hour  , "%d/%m/%Y-%H:%M:%S" ).strftime( "%Y%m%dT%H%M%S" )
   global sequence
   global CRLF
   ics_file.write( f"BEGIN:VEVENT{CRLF}" )
   ics_file.write( f"UID:{debut}-{sequence:06}@fr.thalesgroup.com{CRLF}" )
   ics_file.write( f"DTSTAMP:{debut}{CRLF}" )
   ics_file.write( f"DTSTART:{debut}{CRLF}" )
   ics_file.write( f"DTEND:{fin}{CRLF}" )
   ics_file.write( fold( f"SUMMARY:{objet}" ))
   ics_file.write( fold( f"DESCRIPTION:{description}" ))
   if emplacement:
      ics_file.write( fold( f"LOCATION:{emplacement}" ))
   ndx = organisateur.find('@')
   if ndx > -1:
      dot = organisateur.find('.')
      if dot > -1:
         ndx = dot
      cn = f'{organisateur[0:ndx]}'.capitalize()
      ics_file.write( fold( f'ORGANIZER;CN="{cn}":MAILTO:{organisateur}' ))
   else:
      ics_file.write( fold( f'ORGANIZER;CN="{organisateur}":MAILTO:a.a@a.com' ))
   if "MAHE Aubin" in participants_obligatoires:
      ics_file.write( fold( f'ATTENDEE;ROLE=REQ-PARTICIPANT;CN="{participants_obligatoires}":MAILTO:aubin.mahe@fr.thalesgroup.com' ))
   else:
      ics_file.write( fold( f'ATTENDEE;ROLE=REQ-PARTICIPANT;CN="{participants_obligatoires}":MAILTO:a.a@a.com' ))
   if participants_facultatifs:
      ics_file.write( fold( f'ATTENDEE;ROLE=OPT-PARTICIPANT;CN="{participants_facultatifs}":MAILTO:a.a@a.com' ))
   ics_file.write( f"END:VEVENT{CRLF}" )
   sequence += 1
#
if __name__ == "__main__":
   if len( sys.argv ) != 3:
      print( f'usage: python3 {sys.argv[0]} <Microsoft Office Outlook agenda export CSV file path> <Thunderbird ICS file path>', file = sys.stderr )
   else:
      global CRLF
      CRLF = "\r\n"
      global sequence
      sequence = 1
      with create_header_corrected_CSV( sys.argv[1]       ) as agenda_outlook_file,\
           open(                        sys.argv[2], "wt" ) as ics_file:
         # Lecture de l'entête, mise à jour des clefs d'accès aux champs
         agenda_reader = csv.DictReader( agenda_outlook_file )
         # Ecriture de l'entête
         ics_file.write( f"BEGIN:VCALENDAR{CRLF}" )
         ics_file.write( f"VERSION:2.0{CRLF}" )
         ics_file.write( f"PRODID:-//Aubin.org/NONSGML Windows-Outlook export as ics//EN{CRLF}" )
         # Itération sur chaque entrée
         for row in agenda_reader:
            csv_to_ics( row, ics_file )
         # Ecriture de la fermeture du fichier ics
         ics_file.write( f"END:VCALENDAR{CRLF}" )
