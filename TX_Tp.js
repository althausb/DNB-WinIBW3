function TX_Tp() {
	/*******************************************************************************************************************************

	https://jira.dnb.de/browse/ILT-7211

	16-08-2021  erstellt von: Ali M.
	09-05-2023	bernd	Funktion angepasst und um Unterfunktionen ergänzt, damit sie verteilt werden kann
						neue Unterfunktionen: __getIBWMeldung, __trim, __addTag, _geheZuKat, __dnbDeleteTag
	 ********************************************************************************************************************************/
	 
	function __getIBWMeldung() {
		//Gibt die Meldungen der WinIBW aus
		var msg_ret = "";
		anz = application.activeWindow.messages.count
		for (meld = 0; meld < anz; meld++) {
			msg = application.activeWindow.messages.item(meld);
			msg_ret = msg_ret + msg.text + "\n";
		}
		return msg_ret;
	}
	
	function __trim(string) {
	/*--------------------------------------------------------------------------------------------------------
	Historie:
	2020-09-09 	C. Baumann	als Hilfsfunktion erstellt. Schneidet von einem String Whitespace vorne und hinten ab.
	--------------------------------------------------------------------------------------------------------*/
		if (typeof string == "string")
			return string.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
		else
			return "";
	}
	
	function __geheZuKat(kat, ind, append) {
		/*--------------------------------------------------------------------------------------------------------
		__geheZuKat(kat,ind,append)

		Die Funktion geht in einem Datensatz an die Stelle, an der eine bestimmte neue Kategorie/Indikator der
		Reihenfolge nach eingefügt werden würde. Übergeben wird als Parameter kat die einzufügende Kategorie
		und als ind der Indikator.

		'kat = übergebene einzufügende Kategorie
		'ind = übergebener Indikator
		'append = true -> ans Ende eines vorhandenen Felds (das erste Vorkommen oder, wenn nicht vorhanden, genau ein Feld davor ans Ende), sonst: Anfang des ersten Feldes oder dort, wo es stehen müsste
		'-> append bei noch nicht vorhandenem, einzufügendem Feld immer auf false setzen
		'kat_ind = Wert der Kategorie + Indikator
		'ta_kat = geprüfte Kategorie der TA (Schleife)
		'ta_kat_ind = geprüfte Kategorie der TA + Indikator

		Historie:
		2010-01-09 Stefan Grund		: erstellt
		2010-09-18 Bernd			: Definitionen ergaenzt
		--------------------------------------------------------------------------------------------------------*/

		var ta_kat_ind; // Indikator des Feldes der TA, in dem die richtige Position gesucht wird
		var kat_ind; //
		var ta_kat; //Feld der TA, in dem die richtige Position gesucht wird (pro durgegangener Zeile)
		//kat -> Übergebenes Feld, dessen Postion gesucht werden soll
		//ind -> Übergebener Indikator des übergebenen Feldes, dessen Postion gesucht werden soll

		application.activeWindow.title.startOfBuffer(false);

		do {

			application.activeWindow.title.lineDown(1, false);
			ta_kat = application.activeWindow.title.tag;
			//das gesuchte Feld wurde gefunden, Indikator ist vorhanden
			if (ta_kat == kat && ind != "") {
				kat_ind = parseInt(kat) + parseInt(ind.charCodeAt(0));
				ta_kat_ind = parseInt(ta_kat) + parseInt(application.activeWindow.title.currentField.substr(5, 1).charCodeAt(0));
				//Prüfung: gesuchte kat ungleich Kat der Zeile oder gesuchte Kat + Indikator größer gleich Kat + Indikator der Zeile
				if (ta_kat != kat || ta_kat_ind >= kat_ind) {
					break;
				}
			}
		} while (ta_kat <= kat && ta_kat != ""); //solange Kat der Zeile kleiner gleich gesuchter kat ist und man nicht am Ende eines Datensatzes ist (zu erkennen daran, dass keine Feldbezeichnung vorhanden ist)

		application.activeWindow.title.startOfField(false);

		// Cursor ist jetzt entweder im gesuchten Feld, falls vorhanden, oder im nächsthöheren Feld, falls nicht vorhanden
		if (append == true) {
			//wenn Feld noch nicht vorhanden, ist der Cursor jetzt am Anfang des nächsthöheren Feldes -> muss eins hoch
			if (ta_kat > kat || ta_kat_ind > kat_ind || ta_kat == "") {
				application.activeWindow.title.lineUp(1, false)
			}
			application.activeWindow.title.endOfField(false);
		}
		return application.activeWindow.title.currentField;
	}
	
	function __addTag(tag, content, update) {
		/* Die Funktion geht an die moegliche Position des angegebenen Feldes ("tag") und prueft,
		ob das Feld bereits vorhanden ist.
		Wenn es noch nicht vorhanden ist, wird es in einer neuen Zeile mit Inhalt "content" erzeugt.
		Wenn es bereits vorhanden ist, und das Update ist erwuenscht (=true), wird das bestehende Feld überschrieben. (ACHTUNG: Es werden keine Feldwiederholungen behandelt!)

		Uebersicht der Parameter:
		tag = Die hinzuzufuegende Feldbeschreibung, z.B. "0599"
		content = der im Feld zu ergaenzende Inhalt (das einleitende Leerzeichen wird automatisch ergänzt!) z.B. "f"
		update = true, wenn ein vorhandenes Feld ueberschrieben werden soll, ansonsten false

		2019-09-12 : Marcel Gruss:
		2020-08-31 : Christian Baumann: Einen Bug bereinigt. Zwischen 2 geschuetzte Zeilen muss man zuerst ein '\n'
		schreiben, um dann den Tag einfuegen zu koennen.
		 */

		__geheZuKat(tag, "", true);
		content = " " + content;
		var strTag;
		if (!(strTag = application.activeWindow.title.findTag(tag, 0, true, true, false))) {
			application.activeWindow.title.endOfField(false);
			application.activeWindow.title.insertText("\n")
			application.activeWindow.title.insertText(tag + content);
		} else {
			if (update) {
				application.activeWindow.title.startOfField(false);
				application.activeWindow.title.endOfField(true);
				application.activeWindow.title.deleteSelection();
				application.activeWindow.title.insertText(tag + content);
			}
		}
	}
	
	function __dnbDeleteTag(tag, nurInhalt) {
		/*--------------------------------------------------------------------------------------------------------
		__dnbDeleteTag(tag) boolean

		Die interne Funktion löscht die angegebene Kategorie. Wenn nurInhalt = true, dann wird nur der Inhalt gelöscht, Tag selbst bleibt stehen
		Die Funktion liefert true zurück, wenn die Kategorie gefunden und gelöscht wurde, false, wenn die Kategorie nicht gefunden wurde.

		Historie:
		2010-08-09 Bernd Althaus		: erstellt
		2015-08-31 Stefan Grund 	    : "nurInhalt" ergänzt
		2018-04-11 Stefan Grund			: auch wiederholte Felder werden jetzt gelöscht
		--------------------------------------------------------------------------------------------------------*/
		var zurueck = false;

		y = 0;
		if (application.activeWindow.title) {
			application.activeWindow.title.startOfBuffer(false);
			while (application.activeWindow.title.findTag(tag, y, true, true, false)) {
				if (nurInhalt) {
					application.activeWindow.title.replace(tag + " ");
					y++;
				} else {
					application.activeWindow.title.deleteLine(1);
				}
				zurueck = true;
			}
		}
		return zurueck;

	}
	
	var strScreen = application.activeWindow.getVariable("scr");
	var strTag005 = __trim(application.activeWindow.findTagContent("005", 0, false));
	var str024 = "";
	var x = 0;
	var intwindow = application.activeWindow.windowID;

	if (strScreen == "8A") {

		if (strTag005 == "TX") {
			var y = 0;
			while (str024 = application.activeWindow.findTagContent("024", y, false)) {

				if (str024.indexOf("orcid:") > -1) {

					if (str024.match(/\d{4}-\d{4}-\d{4}-\d{3}[\dX]/)) {

						var orcidNUM;
						if (str024.indexOf("$v") > 0) {
							orcidNUM = str024.substring(str024.indexOf("orcid: ") + 7, str024.indexOf("$v"));
						} else {
							orcidNUM = str024.substring(str024.indexOf("orcid: ") + 7, str024.length);
						}
						//application.messageBox("",str024 + "/" + orcidNUM,"");
						application.activeWindow.command("f stn " + orcidNUM + " and bbg tp*", true);
						var meldung = __getIBWMeldung();
						meldung = __trim(meldung);

						if (meldung == "Nichts gefunden") {
							application.activeWindow.closeWindow();

							application.activeWindow.command("k", false);
							__addTag("005", "Tp", true);
							__geheZuKat("011", false);
							application.activeWindow.title.lineUp(1, false);
							application.activeWindow.title.find("t;f", false, true, false);
							application.activeWindow.title.replace2("f");
							__addTag("040", "$erda", false);
							__addTag("043", "", false);
							__addTag("375", "", false);

							application.activeWindow.title.startOfBuffer(false);

							if (strTag400 = application.activeWindow.title.findTag("400", 0, false, true, false)) {
								if (application.activeWindow.title.find("$vName aus Titel", false, true, false)) {
									application.activeWindow.title.deleteSelection();
								}
							}

							application.activeWindow.title.startOfBuffer(false);
							while (application.activeWindow.title.findTag("667", x, false, false, false)) {

								__dnbDeleteTag("667", false);
								x++;
							}

							__addTag("667", "aus Vorschlagssteuerung$5DE-101", false);

							// in alle 672 Feldern suchen und 2 von denen mit größseste und kleineste Jahr (zwischen $f...$w lassen)
							var x = 0;
							var min = 3000;
							var max = 1;
							while (application.activeWindow.title.findTag("672", x, false, false, false)) {

								var str = application.activeWindow.title.findTag("672", x, false, false, false);
								var strMin = str.substring(str.indexOf("$f") + 2, str.indexOf("$w"));
								var strMax = str.substring(str.indexOf("$f") + 2, str.indexOf("$w"));
								if (strMin < min) {
									min = strMin;
									var str672min = application.activeWindow.title.findTag("672", x, false, false, false);
								}
								if (strMax > max) {
									max = strMax;
									var str672max = application.activeWindow.title.findTag("672", x, false, false, false);
								}
								x++;
							}
							x = 0;
							while (application.activeWindow.title.findTag("672", x, false, false, false)) {
								__dnbDeleteTag("672", false);
								x++;
							}

							if (min != max) {
								__geheZuKat("672", false);
								application.activeWindow.title.insertText("672 " + str672min + "\n");
								application.activeWindow.title.insertText("672 " + str672max + "\n");
							} else {
								__geheZuKat("672", false);
								application.activeWindow.title.insertText("672 " + str672min + "\n");
							}
							__geheZuKat("375", false);
							application.activeWindow.title.lineUp(1, false);
							application.activeWindow.title.endOfField(false);
						} else {
							application.activeWindow.command("s k", false);
							intSetSize = application.activeWindow.getVariable("P3GSZ");
							var txt = "Es besteht bereits " + intSetSize + " Datensatz mit der ORCID ID. Bitte prüfen!";
							application.messageBox("Meldung", txt, "alert-icon");
						}

					} else {
						application.messageBox("Fehler", "Falsches Format der Orcid-Nummer", "error-icon");
					}
				} else {
					y++;
				}
			}

		} else {
			application.messageBox("Fehler", "Kein TX-Datensatz. Bitte prüfen Sie nochmal.", "error-icon");
		}

	} else {
		application.messageBox("Fehler", "Keine Vollanzeige.", "error-icon");
	}
}
