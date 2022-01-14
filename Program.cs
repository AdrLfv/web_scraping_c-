using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Security.Permissions;
using System.Runtime.InteropServices;

namespace Test1WebVisitor
{
	class Program
	{
		static void Main()
		{
			string siret="";
			string siren = "";
			string raison = "";
			string codePostal = "";
			string voie = "";
			string ville = "";
			string effectifs = "";
			bool isRunning = true;
			List<string[]> lignesListe= new List<string[]>();
			int indiceLigne = 0;
			while (isRunning)
            {

				Console.WriteLine("Entrez le lien (par exemple https://www.manageo.fr/entreprises/308197185.html) ou entrez stop");
				string lien = Console.ReadLine();
				if (lien != "stop")
                {

					string chaineInfos = PrincipalFrenchChannels(lien);

					//Console.WriteLine(chaineInfos);
					//Console.ReadLine();
					//Console.Clear();

					siret = IsolationSiret(chaineInfos);
					siren = IsolationSiren(chaineInfos);
					raison = IsolationRaisonSociale(chaineInfos);
					codePostal = IsolationCodePostal(chaineInfos);
					voie = IsolationVoie(chaineInfos);
					ville = IsolationVille(chaineInfos);
					effectifs = IsolationEffectifs(chaineInfos);

					//lignesListe.Add(new Ligne(siret, siren, raison, voie, codePostal, ville));
					string[] tabInfosLigne = new string[8];
					tabInfosLigne[0] = siret;
					tabInfosLigne[1] = siren;
					tabInfosLigne[2] = raison;
					tabInfosLigne[3] = raison;
					tabInfosLigne[4] = codePostal;
					tabInfosLigne[5] = voie;
					tabInfosLigne[6] = ville;
					tabInfosLigne[7] = effectifs;
					lignesListe.Add(tabInfosLigne);

					//Console.WriteLine("Siret : {0}, siren : {1}, raison : {2}, voie : {3}, code postal : {4}, ville : {5}", siret, siren, raison, voie, codePostal, ville);
					
					indiceLigne++;
				}
				else
                {
					isRunning = false;
                }
			}
			AppendFichier(lignesListe);
			/*foreach(string[] element in lignesListe)
            {
				AppendFichier(element[0], element[1], element[2], element[3], element[4], element[5]);
				Console.ReadLine();
			}*/

		}

		static string PrincipalFrenchChannels(string lien)
		{
			Uri uri = new Uri(lien);

			// crée un objet de requête avec l'URI spécifié
			WebRequest request = WebRequest.Create(uri);

			// envoi la requête créee au serveur
			WebResponse response = request.GetResponse();

			// objet de lecture nous permettant de réceptionner le contenu
			// de la réponse du serveur
			StreamReader sr = null;

			try
			{
				// response.GetResponseStream() renvoi un objet
				// de type Stream identifiant le flux de données
				// entre le client (ce programme) et le serveur
				sr = new StreamReader(response.GetResponseStream());
				// lit le flux jusqu'à sa fin
				// (fermeture de la connexion automatique)
				return sr.ReadToEnd();
			}
			catch
			{
				return null;
			}
			finally
			{
				// dans le cas d'une execution sans erreur
				// on prends soin de fermer l'objet de lecture
				// cela facilite le travail du CLR
				if (sr != null)
					sr.Close();

			}

		}
		static string IsolationSiret(string infos)
		{
			int debut = 0;
			debut = infos.IndexOf("data-siret=\"");
			string siret1 = infos.Substring(debut + 12, 14);
			return (siret1);
		}
		static string IsolationSiren(string infos)
		{
			int debut = 0;
			debut = infos.IndexOf("siren=");
			string siren1 = infos.Substring(debut + 7, 9);
			return (siren1);
		}
		static string IsolationRaisonSociale(string infos)
		{
			int debut = 0;
			int taille = 0;
			int fin = 0;
			debut = infos.IndexOf("legalName") + 12;
			fin = infos.IndexOf("\",\"leicode");
			taille = fin - debut;
			string raison1 = (infos.Substring(debut, taille));//.Substring(0, 9);
			return (raison1);
		}
		/*static string IsolationEnseigne(string infos)
		{
			int debut = 0;
			int taille = 0;
			int fin = 0;
			debut = infos.IndexOf("nom-comm") + 19;
			fin = infos.IndexOf("</h2>");
			taille = fin - debut;
			string enseigne1 = (infos.Substring(debut, taille));//.Substring(0, 9);
			return (enseigne1);
		}*/

		/*static string IsolationAdresse(string infos)
		{
			int debut = 0;
			int taille = 0;
			int fin = 0;
			debut = infos.IndexOf("address") + 10;
			fin = infos.IndexOf("\",\"duns\"");
			taille = fin - debut;
			string adresse1 = (infos.Substring(debut, taille));//.Substring(0, 9);


			return (adresse1);
		}*/

		static string IsolationCodePostal(string adresse1)
		{
			int debut = 0;
			debut = adresse1.IndexOf("\"esCp\" value=\"") + 14;
			string codePostal = (adresse1.Substring(debut, 5));
			return (codePostal);
		}

		static string IsolationVoie(string infos)
		{
			int debut = 0;
			int taille = 0;
			int fin = 0;
			debut = infos.IndexOf("text-lowercase \">") + 17;
			fin = infos.IndexOf("</span>,");
			taille = fin - debut;
			string voie1 = (infos.Substring(debut, taille));//.Substring(0, 9);
			return voie1;
		}

		static string IsolationVille(string infos)
		{
			int debut = 0;
			int taille = 0;
			int fin = 0;
			debut = infos.IndexOf("id=\"esVille\" value=\"") + 20;
			fin = infos.IndexOf("esLatitude")-45;
			taille = fin - debut;
			string ville1 = infos.Substring(debut, taille);
			return ville1;
		}

		static string IsolationEffectifs(string infos)
		{
			int debut = 0;
			int taille = 0;
			int fin = 0;
			debut = infos.IndexOf("Tranche d'effectif de l'entreprise") + 74;
			fin = infos.IndexOf("salariés") - 20;
			taille = fin - debut;
			string effectifs1 = infos.Substring(debut, taille);
			return effectifs1;
		}

		static void AppendFichier(List<string[]> lignesListe)
		{

			StreamWriter file = null;
			//string filename = "C:\\Users\\lefev\\Documents\\Classeur1.csv";
			string filename = "C:\\Users\\lefev\\OneDrive\\Documents\\Classeur1.csv";
			try
            {
				file = new StreamWriter(filename);
				foreach (string[] element in lignesListe)
				{
					
					Console.WriteLine(element[0] + ";" + element[1] + ";" + "" + ";" + element[2] + ";" + element[3] + ";" + "" + ";" + "" + ";" + element[4] + ";" + "" + ";" + "" + ";" + element[5] + ";" + element[6] + ";" + element[7]);
					Console.ReadLine();
					string line = /*"\n" +*/ element[0] + ";" + element[1] + ";" + "" + ";" + element[2] + ";" + element[3] + ";" + "" + ";" + "" + ";" + element[4] + ";" + "" + ";" + "" + ";" + element[5] + ";" + element[6] + ";" + element[7];
					file.WriteLine(line);
				}
				
				
			}
			catch (Exception e)
            {
				Console.WriteLine(e.Message);
            }
            finally
            {
				if (file != null)
					file.Close();
            }

			
		}
	}
}