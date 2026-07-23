// ===================================================================
// CONFIG DU CLUB — seul fichier à modifier pour adapter le backend à
// une autre association (barème, effectif, colonnes de la feuille
// Comptes). Le reste du projet ne contient plus aucune donnée propre
// à un club en particulier.
// ===================================================================

const ACTIONS = [
  ["participation mensuelle", 5],
  ["Retard entraînement", 1],
  ["Retard match", 1],
  ["Absence non justifié à l'entrainement", 10],
  ["Oubli de vêtement entraînement", 3],
  ["Oubli de vêtement match", 6],
  ["Absence non justifié au match", 50],
  ["Oubli de chasuble", 2],
  ["Taxer une serviette de douche", 2],
  ["Taxer du savon", 1],
  ["Taxer de la crème", 0.5],
  ["Carton Rouge direct", 7],
  ["2min pour avoir râler", 4],
  ["Carton bleu", 15],
  ["Pas de logo BISCHO pour le déplacement", 5],
  ["Ballon dégueulasse (vraiment !)", 2],
  ["Oubli du ballon (match / entrainement)", 2],
  ["Taxer de l'eau", 1],
  ["Pas présent repas après match domicile", 5],
  ["Nom dans le journal", 2],
  ["Photo dans le journal", 4],
  ["Pire action du match (+ déguisement)", 1],
  ["Meilleure action du match", 2],
  ["Autre (à préciser)", 1],
];

// ATTENTION : ce barème est aussi dupliqué côté frontend (js/config/club-config.js,
// à venir) pour l'affichage. Les deux doivent rester synchronisés manuellement.
const PLAYERS = ["Thomas L.","JM B.","Hugo R.","Yann N.","Arnaud H.","Victor S.","Romain J.","Volodia M.","Victor P.","Mattéo MP.","Gabriel W.","Basile L.","Brahim I.","Maximilien M.","Arthur M.","Damien P.","Alexis W.","Mathieu S.","Nicolas Z.","Robin S."];

// Colonnes de la feuille "Comptes" (une ligne par personne) :
// Nom, Code, Roles, Poste, NomComplet, PhotoURL, Email, PushSubIds
const COL_NOM = 0, COL_CODE = 1, COL_ROLES = 2, COL_POSTE = 3, COL_NOMCOMPLET = 4, COL_PHOTOURL = 5, COL_EMAIL = 6, COL_PUSHSUBIDS = 7;

// Adresse mail de gestion du club, qui reçoit une copie de chaque message envoyé
// depuis la page Support de l'appli.
const CLUB_SUPPORT_EMAIL = "hbcb.gestion@gmail.com";

// Nom du club utilisé comme expéditeur des mails automatiques (rappels de présence, etc.)
const CLUB_NAME = "HBC Bischoffsheim";
