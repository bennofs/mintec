# MINT-EC Zertifikat Generator [![Build Status](https://travis-ci.org/bennofs/mintec-generator.svg?branch=master)](https://travis-ci.org/bennofs/mintec-generator)

Diese Programm dient dazu, aus MINT-EC Zertifikat-Anträgen ein fertiges MINT-EC Zertifikat zu 
erstellen. Es überträgt die Daten aus einer Excel-Tabelle in die Vorlage für das MINT-EC Zertifikat.

*Warnung*: Dieses Programm überprüft nicht, ob die Angaben in den Anträgen korrekt sind (z.B.
ob die angegebenen Wettbewerbsergebnisse auch stimmen!). Die Anträge müssen vorher also manuell
überprüft werden!

## Anleitung

Nach dem Starten des Programms erscheint zunächst ein Dialog zur Auswahl des Ausgabeverzeichnisses.
Das ist das Verzeichnis, in dem die fertigen Zertifikate vom Programm abgelegt werden.

Ist ein Ausgabeverzeichnis ausgewählt, können nun Anträge zur Bearbeitung hinzugefügt werden.
Dies geschieht durch Klicken auf den Button "Hinzufügen", was ein Dialogfenster zur Auswahl
der Anträge öffnet. Hier kann entweder ein Verzeichnis oder ein einzelner Antrag
zum Hinzufügen ausgewählt werden. Bei der Auswahl eines Verzeichnisses werden alle in diesem
Verzeichnis befindlichen Excel-Dateien als Anträge hinzugefügt.

Zum Starten der Generierung muss nun der Button "Zertifikate erstellen" betätigt werden.
Das Programm wandelt dann die Anträge in Zertifikate um, und legt diese im vorher ausgewählten
Ausgabeverzeichnis ab. Festgestellte Probleme werden in der Tabelle in der Spalte "Probleme"
angezeigt.

Die Vorlage zur Erstellung der Zertifikate wird aus der Datei `template.pdf`, die sich im selben
Ordner wie die Jar-Datei des Programms selbst befindet.

## Strukturierung des Quellcodes

Das Projekt verwendet [lombok](https://projectlombok.org/) zur automatischen Generierung von u.a. Gettern und Settern.
Falls eine Java-IDE verwendet wird, muss deshalb möglicherweise ein Plugin für Lombok installiert werden.

Die Hauptklasse des Programms ist in `src/mintec/GUI.java` zu finden. Hier findet die Initialisierung des
Programms statt, welche folgende Schritte umfasst:

* Anzeigen des Dialogs zur Auswahl des Ausgabeverzeichnisses (in main)
* Erstellen einer neuen Instanz der GUI-Klasse, womit das Hauptfenster erzeugt wird (in main)
* Der Konstruktor der GUI-Klasse erzeugt dann die Bedienelemente und lädt außerdem die Datei `template.pdf`
  als Vorlage für die Generierung der MINT-EC Zertifikate.

Der eigentliche Code für die Zertifikaterstellung ist auf die beiden Klassen `MintReader` und `MintWriter` aufgeteilt.
Dabei ist `MintReader` dafür verantwortlich, die Daten aus der Excel-Datei des Antrags auszulesen und zu verarbeiten.
`MintWriter` erzeugt dann aus diesen Daten und der Vorlage ein fertiges Mint-EC Zertifikat, indem die Daten in die
in der Vorlage dafür vorgesehenen Felder eingetragen werden.

Die verbleibenden Klassen sind lediglich dafür zuständig, die Tabelle und den Fortschrittsbalken zu implementieren
und MintReader/MintWriter für alle ausgewählten Dateien auszuführen.

## Erstellen der JAR-Datei aus dem Quellcode
Für das Erstellen verwendet dieses Projekt das ANT-Buildsystem. Das Projekt kann mit folgendem Befehl
erstellt werden:

```bash
$ ant all
```

Alternativ ist im Quellcode auch ein Projekt für die Java-IDE Intellij-IDEA oder Eclipse enthalten.
Da das Projekt Lombok verwendet, muss bei der Verwendung einer Entwicklungsumgebung das Lombok-Plugin
installiert werden.
