program fbRenommeADObjet;

// |===========================================================================|
// | PROGRAM fbRenommeADObjet                                                  |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | Logiciel permetant de renommer des objets dans l'ad                       |
// |___________________________________________________________________________|
// | Ce programme est libre, vous pouvez le redistribuer et ou le modifier     |
// | selon les termes de la Licence Publique Générale GNU publiée par la       |
// | Free Software Foundation .                                                |
// | Ce programme est distribué car potentiellement utile,                     |
// | mais SANS AUCUNE GARANTIE, ni explicite ni implicite,                     |
// | y compris les garanties de commercialisation ou d'adaptation              |
// | dans un but spécifique.                                                   |
// | Reportez-vous à la Licence Publique Générale GNU pour plus de détails.    |
// |                                                                           |
// | anbasso@wanadoo.fr                                                        |
// |___________________________________________________________________________|
// | Versions                                                                  |
// |   1.0.0.0 Création du programme                                           |
// |===========================================================================|


uses
  Forms,
  ufrmMaitre in 'ufrmMaitre.pas' {frmMaitre};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMaitre, frmMaitre);
  Application.Run;
end.
