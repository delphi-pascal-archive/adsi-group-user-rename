unit ufrmMaitre;

// |===========================================================================|
// | unit ufrmMaitre                                                           |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | Unité de la form principal du Logiciel fbrenommeAdObjet                   |
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
// |   1.0.0.0 Création de l'unité                                             |
// |===========================================================================|


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, Spin, ActiveDs_TLB, ufbadsi, activex, ComObj,
  Gauges;

type
  TfrmMaitre = class(TForm)
    gbRecherche: TGroupBox;
    edtOU: TEdit;
    Label1: TLabel;
    gbTypeObjet: TGroupBox;
    cbxUser: TCheckBox;
    cbxGroupe: TCheckBox;
    edtFiltre: TEdit;
    Label2: TLabel;
    cmdRecherche: TButton;
    GroupBox1: TGroupBox;
    strGrid: TStringGrid;
    Label3: TLabel;
    edtMasque: TEdit;
    edtRecherche: TEdit;
    edtRemplacerPar: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    SpinEdit1: TSpinEdit;
    Label6: TLabel;
    cmdSupprimer: TButton;
    cmdRenommer: TButton;
    lblProgression: TLabel;
    SaveDialog1: TSaveDialog;
    cmdSavLog: TButton;
    Gauge1: TGauge;
    procedure cmdSavLogClick(Sender: TObject);
    procedure cmdRenommerClick(Sender: TObject);
    procedure cmdSupprimerClick(Sender: TObject);
    procedure edtRechercheChange(Sender: TObject);
    procedure cmdRechercheClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    procedure initialisestrGrid;
    procedure RedimGrid;
    function rechercheAD : integer;
    function renommerObjet (strOrgNom : string; strMasque : string; intIndice : integer ; strRecherche : string ; strRemplace : string): string;
  public
    { Déclarations publiques }
  end;

var
  frmMaitre: TfrmMaitre;

implementation

{$R *.dfm}
function MatchStrings(source, pattern: String): Boolean;

//  ___________________________________________________________________________
// | Procedure MatchStrings                                                    |
// | _________________________________________________________________________ |
// || Permet de s'avoir si une chaine corespond à un masque de filtre         ||
// ||_________________________________________________________________________||
// || Entrées |  Source : string                                              ||
// ||         |    Chaine à tester                                            ||
// ||         |  pattern : string                                             ||
// ||         |    Masque de filtre                                           ||
// ||_________|_______________________________________________________________||
// || Sorties |  result : boolean                                             ||
// ||         |    vrai si source corespond                                   ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  pSource: Array [0..255] of Char;
  pPattern: Array [0..255] of Char;

  function MatchPattern(element, pattern: PChar): Boolean;
  begin

// ****************************************************************************
// * le masque ="*" c'est ok on sort                                          *
// ****************************************************************************

    if 0 = StrComp(pattern,'*') then
      Result := True
    else

// ****************************************************************************
// * on est à la fin de la chaine source mais pas du masque c'est nok on sort *
// ****************************************************************************

      if (element^ = Chr(0)) and (pattern^ <> Chr(0)) then
        Result := False
      else

// ****************************************************************************
// * on est à la fin de la chaine source et du masque c'est ok on sort        *
// ****************************************************************************

        if element^ = Chr(0) then
          Result := True
        else
        begin
          case pattern^ of

// ****************************************************************************
// * le caractère courant du masque est un "*" on passe au test avec le       *
// * caractère suivant du masque si nok on passe au test avec le caractère    *
// * suivant de source                                                        *
// ****************************************************************************

            '*': if MatchPattern(element,@pattern[1]) then
                   Result := True
                 else
                   Result := MatchPattern(@element[1],pattern);

// ****************************************************************************
// * le caractère courant du masque est un "?" on passe au test avec le       *
// * caractère suivant du masque et du source                                 *
// ****************************************************************************

            '?': Result := MatchPattern(@element[1],@pattern[1]);
          else

// ****************************************************************************
// * le caractère courant du masque corespond au caractère courant de source  *
// * on passe au test le caractère suivant du masque et du source             *
// ****************************************************************************

           if element^ = pattern^ then
             Result := MatchPattern(@element[1],@pattern[1])
           else
            Result := False;
          end;
        end;
  end;

begin
  StrPCopy(pSource,lowercase(source));
  StrPCopy(pPattern,lowercase(pattern));
  Result := MatchPattern(pSource,pPattern);
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Privée                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

Procedure TfrmMaitre.initialisestrGrid ;

//  ___________________________________________________________________________
// | Procedure TfrmMaitre.initialisestrGrid                                    |
// | _________________________________________________________________________ |
// || Permet de réinitialiser la stringgrille                                 ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j : integer;
begin
  for i := 0 to strgrid.RowCount -1 do
    for j := 0 to strgrid.ColCount -1 do
      strgrid.Cells [j,i] :='';
  strgrid.ColCount := 6;
  strgrid.RowCount := 2;
  strgrid.Cells[0,0] := '/!\';
  strgrid.cells[1,0] := 'Nom d''origine';
  strgrid.cells[2,0] := 'Nom Modifié';
  strgrid.cells[3,0] := 'Type';
  strgrid.cells[4,0] := 'Nom LDAP';
  strGrid.cells[5,0] := 'Remarque';
  RedimGrid;
end;


Procedure TfrmMaitre.RedimGrid ;

//  ___________________________________________________________________________
// | Procedure TfrmMaitre.RedimGrid                                            |
// | _________________________________________________________________________ |
// || Permet de redimensionner les colonnes en fonction de leurs contenus     ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j: integer;
  largeur : integer;
begin

// ****************************************************************************
// * Scan de toutes les colonnes                                              *
// ****************************************************************************

  for j := 0 to strgrid.ColCount -1 do
  begin
    largeur := 0;

// ****************************************************************************
// * Scan de toutes les lignes                                                *
// ****************************************************************************

    for i := 0 to strgrid.RowCount -1 do
    begin

// ****************************************************************************
// * Mise à jour en fonction du contenu de la cellule                         *
// ****************************************************************************

      if strgrid.Canvas.TextWidth(strgrid.Cells[j,i])+5 > largeur then
      begin
        largeur := strgrid.Canvas.TextWidth(strgrid.Cells[j,i])+5;
        strgrid.ColWidths [j] := largeur;
      end;
    end;
  end;
  if strGrid.RowCount > 2 then
    if strGrid.Cells[1,strGrid.RowCount-1] ='' then strGrid.RowCount := strGrid.RowCount-1;
end;

function  TfrmMaitre.renommerObjet (strOrgNom : string; strMasque : string; intIndice : integer ; strRecherche : string ; strRemplace : string): string;
//  ___________________________________________________________________________
// | function  tfrmMaitre.renommerObjet                                        |
// | _________________________________________________________________________ |
// || Permet de renommer un objet en fonction d'un masque                     ||
// ||_________________________________________________________________________||
// || Entrées | strOrgNom : string                                            ||
// ||         |   Nom d'origine                                               ||
// ||         | strMasque : string                                            ||
// ||         |   Masque de renommage '*' est remplacé par le nom d'origine   ||
// ||         |   et '#' par intindice                                        ||
// ||         | intIndice : integer                                           ||
// ||         |   numero à afficher à la place des '#"                        ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   nom modifié                                                 ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|
var
  strindice : string;
  i,j: integer;
  strTemp : string;
begin

// ****************************************************************************
// * Remplacement d'une partie par une autre                                  *
// ****************************************************************************

  strTemp := StringReplace(strOrgNom  ,strRecherche ,strRemplace ,[rfReplaceAll,rfIgnoreCase]);

// ****************************************************************************
// * Remplacement des # par l'indice                                          *
// ****************************************************************************

  strindice :='0' + inttostr(intIndice );
  j := length(strindice );
  for i := length(strMasque) downto 1 do
    if strMasque[i] ='#' then
    begin
      strMasque[i]:= strindice[j];
      j := j -1 ;
      if j <1 then j := 1;
    end else j := length(strindice );

// ****************************************************************************
// * Remplacement des * par le nom d'origine                                  *
// ****************************************************************************

  result := StringReplace(strMasque ,'*',strTemp,[rfReplaceAll]);

// ****************************************************************************
// * Si le resultat est vide on garde le nom d'origine                        *
// ****************************************************************************

  if result ='' then result := strOrgNom;
end;

function TfrmMaitre.rechercheAD :integer ;

//  ___________________________________________________________________________
// | function  TfrmMaitre.rechercheAD                                          |
// | _________________________________________________________________________ |
// || Permet de recherhcer des objets dans l'AD                               ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties | result : integer                                              ||
// ||         |   Nombre d'objets trouvés -1 en cas d'erreur                  ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  search : IDirectorySearch;
  p : array[0..0] of PWideChar;
  ptrResult : THandle;
  col : ads_search_column ;
  hr : HResult;
  opt : array[0..0] of ads_searchpref_info;
  dwErr : DWord;
  szErr : array[0..255] of WideCHar;
  szName : array[0..255] of WideChar;
  strtemp : string;
  adobjet : IADs ;
  strrecherche : string;
  strNom : string;
begin
  result := 0;
  AdsGetObject(edtOU.Text , IDirectorySearch, search);
  p[0] := StringToOleStr('adspath');
  opt[0].dwSearchPref :=ADS_SEARCHPREF_PAGESIZE;
  opt[0].vValue.dwType := ADSTYPE_INTEGER;
  opt[0].vValue.Integer := 100;
  hr := search.SetSearchPreference(@opt[0],1);
  if (hr <> 0) then
  begin
    ADsGetLastError(dwErr, @szErr[0], 254, @szName[0], 254);
    ShowMessage(WideCharToString(szErr));
    Exit;
  end;

// ****************************************************************************
// * Construction chaine de recherche                                         *
// ****************************************************************************

  strrecherche := '';
  if cbxUser.Checked then strrecherche := strrecherche + '(objectCategory=user)';
  if cbxGroupe.Checked then strrecherche := strrecherche  + '(objectCategory=group)';
  strrecherche := '(|'+ strrecherche +')';
  search.ExecuteSearch(strrecherche,@p[0], 1, ptrResult);
  hr := search.GetNextRow(ptrResult);
  while (hr <> S_ADS_NOMORE_ROWS) do
  begin
    strtemp := '';
    hr := search.GetColumn(ptrResult, p[0],col);
    if Succeeded(hr) then
    begin
      if col.pADsValues <> nil then
      begin
        try

// ****************************************************************************
// * Pour chaque objet trouvé si le nom correspond aux critères de filtre on  *
// * l'ajoute                                                                 *
// ****************************************************************************

          strtemp := col.pAdsvalues^.ClassName;
          AdsGetObject(strtemp,IADs,adobjet);
          strNom := copy(adobjet.Name,4,length(adobjet.Name)-3);
          if MatchStrings( strNom , edtFiltre.Text ) then
          begin
            strGrid.RowCount := strGrid.RowCount +1;
            if strgrid.RowCount mod 10 = 0 then
            begin
              lblProgression.Caption := inttostr(strgrid.RowCount-2);
              application.ProcessMessages ;
            end;
            strGrid.Cells[4,strGrid.RowCount -2] := strtemp;
            strGrid.Cells[3,strGrid.RowCount -2] := adobjet.Class_ ;
            strgrid.Cells[1,strGrid.RowCount -2] := strNom;
          end;
        except
        end;
      end;
    end;
    search.FreeColumn(col);
    Hr := search.GetNextRow(ptrResult);
  end;
  lblProgression.Caption := 'Recherche terminée : ' + inttostr(strgrid.RowCount-2);
  search.CloseSearchHandle(ptrResult)
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Evénementiel                                                       #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmMaitre.FormCreate(Sender: TObject);

//  ___________________________________________________________________________
// | function  TfrmMaitre.FormCreate                                           |
// | _________________________________________________________________________ |
// || Permet de recherhcer des objets dans l'AD                               ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant la procedure                                 ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  initialisestrGrid ;
end;

procedure TfrmMaitre.cmdRechercheClick(Sender: TObject);

//  ___________________________________________________________________________
// | function  TfrmMaitre.cmdRechercheClick                                    |
// | _________________________________________________________________________ |
// || Permet de recherhcer des objets dans l'AD                               ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant la procedure                                 ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if edtou.Text ='' then exit;
  initialisestrGrid ;
  rechercheAD ;
  edtRechercheChange(nil);
  RedimGrid ;
  cmdRenommer.Enabled := true;
end;

procedure TfrmMaitre.edtRechercheChange(Sender: TObject);

//  ___________________________________________________________________________
// | function  TfrmMaitre.edtRechercheChange                                   |
// | _________________________________________________________________________ |
// || Permet d'actualiser le resultat du renommage                            ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant la procedure                                 ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,compteur : integer;
begin
  compteur := SpinEdit1.Value ;
  for i := 1 to strGrid.RowCount - 1 do
  begin
    strGrid.Cells[2,i] := renommerObjet ( strGrid.Cells [1,i],edtMasque.text,compteur , edtRecherche.Text ,edtRemplacerPar.Text );
    inc(compteur,1);
  end;
  RedimGrid ;
end;

procedure TfrmMaitre.cmdSupprimerClick(Sender: TObject);

//  ___________________________________________________________________________
// | function  TfrmMaitre.cmdSupprimerClick                                    |
// | _________________________________________________________________________ |
// || Permet de supprimer la ligne courrante de la grille                     ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant la procedure                                 ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var i,j : integer;
begin
  if strGrid.Row > 0 then
    if strGrid.cells[1,strGrid.Row] <> '' then
    begin
      for i := strGrid.Row to strGrid.RowCount -2 do
        for j := 0 to strGrid.ColCount -1 do
        begin
          strgrid.Cells[j,i] := strGrid.Cells[j,i+1];
          strGrid.Cells[j,i+1] := '';
        end;
      if strGrid.RowCount > 2 then strGrid.RowCount :=  strGrid.RowCount-1;
    end;
end;

procedure TfrmMaitre.cmdRenommerClick(Sender: TObject);

//  ___________________________________________________________________________
// | function  TfrmMaitre.cmdRenommerClick                                     |
// | _________________________________________________________________________ |
// || Permet de lancer le renommage des objets                                ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant la procedure                                 ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
  adUser : IADsUser;
  adGrp : IADsGroup ;
  adcont : IADsContainer ;

begin
  cmdRenommer.Enabled := false;
  Gauge1.Progress := 0;
  gauge1.MaxValue := strGrid.RowCount -1;
  for i := 1 to strGrid.RowCount -1 do
  begin
    if strGrid.Cells[2,i] <> strGrid.Cells[1,i] then
    begin

// ****************************************************************************
// * Renommage d'un groupe                                                    *
// ****************************************************************************

      if strGrid.Cells [3,i] = 'group' then
      try

// ****************************************************************************
// * Modification du nom visible                                              *
// ****************************************************************************

        AdsGetObject(strGrid.cells[4,i],IADsGroup,adGrp ) ;
        adgrp.Put('samAccountName',strGrid.Cells[2,i]);
        adGrp.SetInfo;

// ****************************************************************************
// * Modification du nom en passant par la fonction MoveHere de l'OU contenant*
// * l'objet ceci permet soit le déplacement de l'objet ou son renommage      *
// ****************************************************************************

        AdsGetObject(adGrp.Parent, IADsContainer,adcont) ;
        adcont.MoveHere(strgrid.cells[4,i],'CN='+ strGrid.Cells[2,i]);
        strGrid.Cells [0,i] := 'OK';
      except
      on e : exception do
        begin
          strGrid.Cells [0,i] := 'NOK';
          strGrid.cells [5,i] := e.Message ;
        end;
      end;

// ****************************************************************************
// * Renommage d'un utilisateur                                               *
// ****************************************************************************

      if strGrid.Cells [3,i] = 'user' then
      try

// ****************************************************************************
// * Modification du nom visible                                              *
// ****************************************************************************

        AdsGetObject(strGrid.cells[4,i], IADsUser ,aduser);
        aduser.Put('displayName',strGrid.Cells[2,i]);
        aduser.SetInfo;

// ****************************************************************************
// * Modification du nom en passant par la fonction MoveHere de l'OU contenant*
// * l'objet ceci permet soit le déplacement de l'objet ou son renommage      *
// ****************************************************************************

        AdsGetObject(aduser.Parent,IADsContainer,adcont) ;
        adcont.MoveHere(strgrid.cells[4,i],'CN='+ strGrid.Cells[2,i]);
        strGrid.Cells [0,i] := 'OK';
      except
      on e : exception do
        begin
          strGrid.Cells [0,i] := 'NOK';
          strGrid.cells [5,i] := e.Message ;
        end;
      end;
    end;
    gauge1.Progress := i-1;
  end;
  RedimGrid;
end;

procedure TfrmMaitre.cmdSavLogClick(Sender: TObject);

//  ___________________________________________________________________________
// | function  TfrmMaitre.cmdSavLogClick                                       |
// | _________________________________________________________________________ |
// || Permet de sauvegarder le contenu de la grille dans un fichier text      ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant la procedure                                 ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  fichier : textfile;
  i,j : integer;
  strtemp : string;
begin
  if SaveDialog1.Execute then
  begin

// ****************************************************************************
// * Création du fichier                                                      *
// ****************************************************************************

    assignfile(fichier,SaveDialog1.FileName );
    rewrite(fichier );


// ****************************************************************************
// * On parcoure toutes les lignes de la grille                               *
// ****************************************************************************

    for j := 0 to strGrid.RowCount -1 do
    begin

// ****************************************************************************
// * Création de la chaine avec toutes les colonnes séparées par un ";"       *
// ****************************************************************************

      strtemp := '';
      for i := 0 to strGrid.ColCount -1 do
        strtemp := strtemp + strGrid.Cells[i,j] + ';';

// ****************************************************************************
// * Ecriture de la chaine dans le fichier                                    *
// ****************************************************************************

      writeln(fichier,strtemp);
    end;

// ****************************************************************************
// * Fermeture du fichier                                                     *
// ****************************************************************************

    CloseFile(fichier);
  end;
end;

end.
