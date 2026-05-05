import { useState, useEffect, useMemo } from "react";
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const DATA_ABE = [
  {id:"P27/ABE",num:1,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"27",nom:"RANDRIANARIMASY Radosoa",cin:"",contact:"034 03 846 09",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:5360000.0,solde:34640000.0,mensualites_ecoulees:8.0,mensualites_restantes:52.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P63/ABE",num:2,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"63",nom:"RAHARISON Tsiriaina Miarana",cin:"",contact:"034 05 536 05",correspondance:"344322067",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P47/ABE",num:3,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"47",nom:"RAHELIARINIVO Jillà Onitriniaina",cin:"",contact:"034 02 002 86",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:4690000.0,solde:35310000.0,mensualites_ecoulees:7.0,mensualites_restantes:53.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P29/ABE",num:4,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"29",nom:"HAJARISAONA Elima",cin:"",contact:"034 91 438 69",correspondance:"345406857",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P14/ABE",num:5,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"14",nom:"RAZAFINDRAVELO Pascaline",cin:"",contact:"034 39 012 18",correspondance:"340118887",situation:"légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:3360000.0,solde:36640000.0,mensualites_ecoulees:5.01,mensualites_restantes:54.99,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P50/ABE",num:6,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"50",nom:"RAMANANARIVO Nivo Mirelle",cin:"",contact:"034 01 838 51",correspondance:"rasoanaivoboris@gmail.com",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P51/ABE",num:7,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"51",nom:"ANDRIAMARO Daudet Firmin",cin:"",contact:"034 79 733 46",correspondance:"346478784",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P10/ABE",num:8,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"10",nom:"RANDRIATSALAMA Hajarisoa",cin:"",contact:"034 24 423 29",correspondance:"340752804",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:4690000.0,solde:35310000.0,mensualites_ecoulees:7.0,mensualites_restantes:53.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P61/ABE",num:9,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"61",nom:"HANTARIVELO Karl Florence",cin:"",contact:"038 94 868 45",correspondance:"346474122",situation:"signé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P59/ABE",num:10,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"59",nom:"ANDRIATSARAFARA Hanitriniony",cin:"",contact:"034 01 502 07",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P15/ABE",num:11,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"15",nom:"RAKOTOZAFY Nomenjanahary",cin:"",contact:"034 01 805 55",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P68/ABE",num:12,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"68",nom:"RAIVOARISOA Monique",cin:"",contact:"034 08 727 52",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P11/ABE",num:13,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"11",nom:"ANDRIAMANALINA Mamisantatriniaina",cin:"",contact:"034 39 136 02",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P33/ABE",num:14,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"33",nom:"RAVELOSON Jacky Etienne",cin:"",contact:"034 60 759 34/033 32 948 50",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P77/ABE",num:15,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"77",nom:"RAKOTOARISOA Lovaniaina",cin:"",contact:"034 28 036 58",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:5360000.0,solde:34640000.0,mensualites_ecoulees:8.0,mensualites_restantes:52.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P8/ABE",num:16,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"8",nom:"RAKOTONANDRASANA Herinirina",cin:"",contact:"034 03 700 09",correspondance:"340552434",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:4020000.0,solde:35980000.0,mensualites_ecoulees:6.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P43/ABE",num:17,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"43",nom:"RABARISON Holiarisoa  Vonihanitriniaina",cin:"",contact:"034 20 066 05",correspondance:"",situation:"légalisé",date_signature:"2025-07-17",date_fin:"2030-07-17",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:1340000.0,solde:38660000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P71/ABE",num:18,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"71",nom:"RAHELINIRINA Felanarisoa Modestine",cin:"",contact:"033 23 429 25",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P9/ABE",num:19,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"9",nom:"RATSIMBAZAFY Hantanirina Sylvia",cin:"",contact:"034 05 549 28",correspondance:"342532902",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P24/ABE",num:20,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"24",nom:"HOBINJATOVO Tokiniaina",cin:"",contact:"034 14 046 48",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:1340000.0,solde:38660000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P56/ABE",num:21,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"56",nom:"RAKOTONDRAMASY Linah Augustine",cin:"",contact:"034 87 154 17",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P66/ABE",num:22,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"66",nom:"RAKOTOZAFY Jean Michel",cin:"",contact:"034 46 683 34",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P31/ABE",num:23,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"31",nom:"RAKOTONIRINA Prosper",cin:"",contact:"034 74 868 75",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P60/ABE",num:25,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"60",nom:"ANDRIANJAFY Nadia Elisa",cin:"",contact:"034 42 414 58",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P69/ABE",num:26,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"69",nom:"JEDIDIA Ranivo Harinoely",cin:"",contact:"034 01 022 34",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P67/ABE",num:27,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"67",nom:"TIARAY Mirantsoa",cin:"",contact:"034 13 379 13",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P35/ABE",num:28,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"35",nom:"RANDRIAMANARIVO Herilala Jean Luck",cin:"",contact:"034 71 712 36",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P19/ABE",num:29,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"19",nom:"RAKOTONARIVO Fidy Mpanjato",cin:"",contact:"034 02 311 81",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:4020000.0,solde:35980000.0,mensualites_ecoulees:6.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P22/ABE",num:30,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"22",nom:"RANDRIANARIVONY Lovasoa",cin:"",contact:"038 01 000 34",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P23/ABE",num:31,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"23",nom:"RIVOSOA Malalatiana Sylvie",cin:"",contact:"034 53 745 06",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:4020000.0,solde:35980000.0,mensualites_ecoulees:6.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P12/ABE",num:32,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"12",nom:"RAKOTOBE Fanomezanjo Tolotra",cin:"",contact:"034 11 732 90",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P2/ABE",num:33,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"2",nom:"RANDRIANASOLO Josea Todisoa",cin:"",contact:"034 91 923 73",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P6/ABE",num:35,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"6",nom:"RANDIMBY Irintsoa Ainjatovo",cin:"",contact:"034 87 303 17",correspondance:"",situation:"signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P7/ABE",num:36,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"7",nom:"RIVOLALAINA Paul Jocelyn",cin:"",contact:"034 49 046 71",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:4020000.0,solde:35980000.0,mensualites_ecoulees:6.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P62/ABE",num:37,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"62",nom:"RAJAONARIVELO Rindra Niriana",cin:"",contact:"034 71 488 63",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P28/ABE",num:38,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"28",nom:"RAKOTOMANDIMBY Rijanirina",cin:"",contact:"034 54 711 39",correspondance:"",situation:"légalisé",date_signature:"2025-06-01",date_fin:"2030-06-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P13/ABE",num:39,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"13",nom:"RAZAFIMAHEFA Miandrisoa Lovasoa",cin:"",contact:"034 80 959 35",correspondance:"",situation:"légalisé",date_signature:"2025-09-02",date_fin:"2030-09-02",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P17/ABE",num:40,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"17",nom:"RASOAFARA Miora Hanitriniaina",cin:"",contact:"034 43 451 73",correspondance:"",situation:"légalisé",date_signature:"2025-08-29",date_fin:"2030-08-29",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P18/ABE",num:41,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"18",nom:"RAKOTONDRAZAFY Ndrianja Harijaona",cin:"",contact:"032 42 150 62",correspondance:"",situation:"légalisé",date_signature:"2025-08-01",date_fin:"2030-08-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:3280000.0,solde:36720000.0,mensualites_ecoulees:4.9,mensualites_restantes:55.1,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P20/ABE",num:42,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"20",nom:"RAZANALISOA Mbolatahina",cin:"",contact:"034 05 548 66",correspondance:"",situation:"légalisé",date_signature:"2025-08-01",date_fin:"2030-08-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:1340000.0,solde:38660000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P25/ABE",num:43,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"25",nom:"RAMIARINJAONA Lovanantenaina Sylvia",cin:"",contact:"034 08 732 29",correspondance:"",situation:"signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P26/ABE",num:44,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"26",nom:"RAOLIARISOA Charlotte",cin:"",contact:"034 98 122 71",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P37/ABE",num:45,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"37",nom:"JAOTIANA Igore",cin:"",contact:"034 31 450 71",correspondance:"",situation:"légalisé",date_signature:"2025-08-25",date_fin:"2030-08-25",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P42/ABE",num:46,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"42",nom:"RAMANJATONIRINA  Vao Voahanginiaina Faramalala",cin:"",contact:"032 25 078 99",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P44/ABE",num:47,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"44",nom:"RABERANTO Jean William Seth",cin:"",contact:"033 13 137 10",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P52/ABE",num:48,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"52",nom:"RAKOTONDRAIBE Fahadimy Odilon",cin:"",contact:"034 27 983 20",correspondance:"",situation:"légalisé",date_signature:"2025-08-26",date_fin:"2030-08-26",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P57/ABE",num:50,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"57",nom:"RANDRIANARISON Vonitiana Rolland",cin:"",contact:"034 01 754 71",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P64/ABE",num:51,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"64",nom:"VOAVY CLARISSE Vololoniaina",cin:"",contact:"034 47 421 56",correspondance:"",situation:"légalisé",date_signature:"2025-09-23",date_fin:"2030-09-23",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P65/ABE",num:52,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"65",nom:"ANDRIAMANANIARY Fety Mahery Zo",cin:"",contact:"034 12  648 99",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P76/ABE",num:53,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"76",nom:"HARIMANANA Solofonirina",cin:"",contact:"034 60 163 63",correspondance:"",situation:"Signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P80/ABE",num:54,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"80",nom:"REMAMY Ndimbiarijaona Randzavola Mahefa",cin:"",contact:"034 38 502 61",correspondance:"",situation:"légalisé",date_signature:"2025-07-20",date_fin:"2030-07-20",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P49/ABE",num:55,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"49",nom:"NJANDRY Arly Frédiarisoa",cin:"",contact:"034 26 500 14",correspondance:"",situation:"légalisé",date_signature:"2025-09-04",date_fin:"2030-09-04",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P36/ABE",num:57,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"36",nom:"ANDRIANIAINA TOKY MIONONA",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-08",date_fin:"2030-09-08",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"miandry trano vita, mbola mandoa hofatrano ao @ lot 1116 E 110 Tsitamaso",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P48/ABE",num:58,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"48",nom:"RAZAFINDRASOA NOROTIANA FRANCOISE",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"miandry trano vita fa mandoa hofatrano sady ampiditra mpianatra au Lot 304E 284 Andafiantsimo Star",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P78/ABE",num:59,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"78",nom:"MAHERITIANA LUCIEN ANTHONIO",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"miandry trano vita satria mbola mandoa hofatrano au lot 304E284 Andafiantsimo Star",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P45/ABE",num:60,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"45",nom:"AMBININTSOA MALALATIANA NATHALIE",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-04",date_fin:"2030-09-04",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"mbola miandry tirage logement vao hanapakevitra hoe hanao contrat ou pas",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P73/ABE",num:61,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"73",nom:"HOBINANDRAINA MANOA ZO HERINIAINA",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"injoignable",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P38/ABE",num:62,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"38",nom:"MBOLATINA ALIE ZYNA",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-08-27",date_fin:"2030-08-27",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:1340000.0,solde:38660000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"attente retour jsq fin Juin",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P4/ABE",num:63,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"4",nom:"RAJAONERA MAHERY ANDRIAMANALINA",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-01",date_fin:"2030-09-01",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"tsy mbola manana lgmt/Analogh miova2 foana/anaty contrat hoe mahazo lakilé fa manao contrat Lettre de relance notifiée - 7 jours à partir du 01/08/2025- Attente confirmation de signature",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P3/ABE",num:64,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"3",nom:"ANDRIAMANIVONIONY BAKOLIMANALINIRINA AUGUSTINE",cin:"",contact:"",correspondance:"",situation:"signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"Confirmation le 05/08",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P75/ABE",num:65,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"75",nom:"HERIHAJANIAVO AUGUSTIN MARTIAL",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"miandry ny avancement début sept. Afaka manao contrat na mbola mandoa hofatrano ary rhf mvoka ny mandatement",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P40/ABE",num:66,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"40",nom:"RANDRIAMIARISON Zilga",cin:"",contact:"034 68 458 02",correspondance:"",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P58/ABE",num:73,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"58",nom:"RANDRIANARISON Niandritiana Rolland",cin:"",contact:"",correspondance:"",situation:"Signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"mbola iresaka @ bq ny rsk farany dia zao injoignable",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P/ABE",num:76,sigle:"ABE",localite:"Vatofotsy Antsirabe",num_lgt:"",nom:"RANAIVOSAMIMANANA Tolotra Stéphane",cin:"",contact:"",correspondance:"",situation:"Signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
];
const DATA_IMT = [
  {id:"P81/IMT",num:1,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"81",nom:"ANDRIANALIHAJA Miary Rakotomalala",cin:"501011004306.0",contact:"344322067",correspondance:"346978037.0",situation:"Légalisé",date_signature:"2024-09-05",date_fin:"2029-09-05",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:9380000.0,solde:30620000.0,mensualites_ecoulees:14.0,mensualites_restantes:47.0,note:"fin sept",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P110/IMT",num:5,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"110",nom:"RANAIVOSOA Boris",cin:"201211155287.0",contact:"344955557",correspondance:"",situation:"légalisé",date_signature:"2024-09-19",date_fin:"2029-09-19",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:7370000.0,solde:32630000.0,mensualites_ecoulees:11.0,mensualites_restantes:48.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P287/IMT",num:6,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"287",nom:"TAHINARISOA Zanamida Cynthia",cin:"515012043236.0",contact:"346478784",correspondance:"",situation:"légalisé",date_signature:"2024-09-20",date_fin:"2029-09-20",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:10720000.0,solde:29280000.0,mensualites_ecoulees:16.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P9/IMT",num:7,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"9",nom:"RANDRIAMAMISOA Mahandriarison Nitsinjo Fitahiana",cin:"106302030760.0",contact:"340752804",correspondance:"",situation:"signé",date_signature:"2024-08-28",date_fin:"2029-08-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:7370000.0,solde:32630000.0,mensualites_ecoulees:11.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P124/IMT",num:8,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"124",nom:"RAHERINIRINA Rollande",cin:"102032008264.0",contact:"346474122",correspondance:"341674178.0",situation:"légalisé",date_signature:"2024-08-28",date_fin:"2029-08-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8710000.0,solde:31290000.0,mensualites_ecoulees:13.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P111/IMT",num:9,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"111",nom:"RANDRIATSARAFARA Tanjona Mirindra",cin:"117032010013.0",contact:"333753695",correspondance:"",situation:"Légalisé",date_signature:"2024-08-28",date_fin:"2029-08-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"injoignable",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P41/IMT",num:10,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"41",nom:"RASOHARININAHARINALA Ursule",cin:"105992006316.0",contact:"341097475/0340034667",correspondance:"327462948.0",situation:"signé",date_signature:"2024-08-28",date_fin:"2029-08-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8040000.0,solde:31960000.0,mensualites_ecoulees:12.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P194/IMT",num:11,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"194",nom:"RANDRIATSALAMA Lolona",cin:"117052003172.0",contact:"343568371",correspondance:"340681423.0",situation:"légalisé",date_signature:"2024-08-29",date_fin:"2029-08-29",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6700000.0,solde:33300000.0,mensualites_ecoulees:10.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P212/IMT",num:14,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"212",nom:"RAMILISONINA Lucie Olga",cin:"101222044077.0",contact:"340937634",correspondance:"",situation:"Légalisé",date_signature:"2024-09-02",date_fin:"2029-09-02",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8710000.0,solde:31290000.0,mensualites_ecoulees:13.0,mensualites_restantes:47.0,note:"annulé",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P62/IMT",num:16,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"62",nom:"RASOARINERA Nandrasana Toditanjona",cin:"301092026566.0",contact:"345423457",correspondance:"342985210.0",situation:"signé",date_signature:"2024-11-20",date_fin:"2029-11-20",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"injoignable",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P116/IMT",num:18,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"116",nom:"BAKOLY Notahiana",cin:"108072001617.0",contact:"340199747",correspondance:"349622926.0",situation:"Légalisé",date_signature:"2024-09-27",date_fin:"2029-09-27",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P186/IMT",num:19,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"186",nom:"RANAIVOSON Voanginirina",cin:"",contact:"342811133",correspondance:"",situation:"légalisé",date_signature:"2025-03-06",date_fin:"2030-03-06",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:53.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P141/IMT",num:20,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"141",nom:"RASOANOMENJANAHARY Anjarasoa Maharavo",cin:"101232108508.0",contact:"349432105",correspondance:"349691659.0",situation:"Légalisé",date_signature:"2024-09-24",date_fin:"2029-09-24",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8040000.0,solde:31960000.0,mensualites_ecoulees:12.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P276/IMT",num:21,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"276",nom:"RAMANAMPISOA Nivoharivela",cin:"210152005822.0",contact:"341833312/348053807",correspondance:"",situation:"légalisé",date_signature:"2024-09-24",date_fin:"2029-09-24",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P100/IMT",num:22,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"100",nom:"RANDRIAMAMPIANINA Faniry Ny Aina",cin:"101251169225.0",contact:"0325106761 -  0331566618",correspondance:"",situation:"légalisé",date_signature:"2024-09-20",date_fin:"2029-09-20",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8040000.0,solde:31960000.0,mensualites_ecoulees:12.0,mensualites_restantes:48.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P203/IMT",num:23,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"203",nom:"RAKOTOARISON Andolalao",cin:"101982064411.0",contact:"344429663",correspondance:"",situation:"légalisé",date_signature:"2024-09-27",date_fin:"2027-09-27",prix_logement:40000000,mensualite:1120000,nb_mensualites:36,loyer_paye:17920000.0,solde:22080000.0,mensualites_ecoulees:16.0,mensualites_restantes:24.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P163/IMT",num:24,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"163",nom:"RANAIVOZAKA Jean Tonny",cin:"117311008895.0",contact:"343631304",correspondance:"",situation:"Légalisé",date_signature:"2024-10-02",date_fin:"2029-10-02",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P308/IMT",num:25,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"308",nom:"RAJOELISON Hanitrasoa",cin:"101232108345.0",contact:"346978037",correspondance:"",situation:"Légalisé",date_signature:"2024-10-09",date_fin:"2029-10-09",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2680000.0,solde:37320000.0,mensualites_ecoulees:4.0,mensualites_restantes:51.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P147/IMT",num:26,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"147",nom:"RAHARIMALALASOA Marie Audolphine",cin:"301092012606.0",contact:"347412660",correspondance:"",situation:"légalisé",date_signature:"2024-10-18",date_fin:"2029-10-18",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P232/IMT",num:27,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"232",nom:"MAMY ANDRIANAVONISON Carole Stéphanie",cin:"112992011646.0",contact:"347152626",correspondance:"",situation:"Légalisé",date_signature:"2024-10-18",date_fin:"2029-10-18",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P208/IMT",num:28,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"208",nom:"MIRANA Prisca Verone Luce",cin:"501052004459.0",contact:"327531601",correspondance:"",situation:"Légalisé",date_signature:"2024-10-31",date_fin:"2029-10-31",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8710000.0,solde:31290000.0,mensualites_ecoulees:13.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P139/IMT",num:29,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"139",nom:"RAFARASOA Tina Domoina Cassidy",cin:"101212158676.0",contact:"346060397",correspondance:"",situation:"Légalisé",date_signature:"2024-11-27",date_fin:"2029-11-27",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P315/IMT",num:30,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"315",nom:"RANDRIANIRINA Rondro Lalaina Sylvia",cin:"101222052890.0",contact:"340133365",correspondance:"",situation:"légalisé",date_signature:"2024-12-19",date_fin:"2029-12-19",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P153/IMT",num:31,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"153",nom:"RAKOTONJANAHARY Alfred",cin:"117351005262.0",contact:"346677330",correspondance:"",situation:"légalisé",date_signature:"2025-01-21",date_fin:"2030-01-21",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P10/IMT",num:32,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"10",nom:"RASOANAIVO Francis Odilon",cin:"101241038832.0",contact:"341674178",correspondance:"",situation:"légalisé",date_signature:"2025-02-06",date_fin:"2030-02-06",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:10050000.0,solde:29950000.0,mensualites_ecoulees:15.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P39/IMT",num:33,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"39",nom:"ANDRIANARILALA Sedra Famantanantsoa",cin:"105011012133.0",contact:"348865529",correspondance:"",situation:"légalisé",date_signature:"2025-02-14",date_fin:"2030-02-14",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:1340000.0,solde:38660000.0,mensualites_ecoulees:2.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P145/IMT",num:34,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"145",nom:"RASOAHARIMALALA Florine",cin:"301032005742.0",contact:"342597939",correspondance:"",situation:"légalisé",date_signature:"2025-03-05",date_fin:"2030-03-05",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:10050000.0,solde:29950000.0,mensualites_ecoulees:15.0,mensualites_restantes:47.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P230/IMT",num:40,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"230",nom:"RANDIMBIARISON Naina Mamisoa",cin:"",contact:"340228014",correspondance:"",situation:"légalisé",date_signature:"2025-03-18",date_fin:"2030-03-18",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:10050000.0,solde:29950000.0,mensualites_ecoulees:15.0,mensualites_restantes:48.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P54/IMT",num:41,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"54",nom:"KOTOMARO Andrianomena Daniela",cin:"",contact:"342985210",correspondance:"328018230.0",situation:"légalisé",date_signature:"2025-03-07",date_fin:"2030-03-07",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:7370000.0,solde:32630000.0,mensualites_ecoulees:11.0,mensualites_restantes:48.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P142/IMT",num:43,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"142",nom:"RAKOTONDRAZAFY Hariniaina H",cin:"",contact:"349622926",correspondance:"",situation:"signé",date_signature:"2025-03-12",date_fin:"2030-03-12",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6700000.0,solde:33300000.0,mensualites_ecoulees:10.0,mensualites_restantes:48.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P202/IMT",num:44,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"202",nom:"BENANDRASANA Aimé",cin:"",contact:"338734380/034 27 877 89",correspondance:"",situation:"légalisé",date_signature:"2025-03-07",date_fin:"2030-03-07",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8040000.0,solde:31960000.0,mensualites_ecoulees:12.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P58/IMT",num:45,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"58",nom:"RASOANAIVO Voahirana Emma Julia",cin:"",contact:"349691659",correspondance:"",situation:"signé",date_signature:"2025-06-12",date_fin:"2030-06-12",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8710000.0,solde:31290000.0,mensualites_ecoulees:13.0,mensualites_restantes:48.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P213/IMT",num:62,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"213",nom:"HARINALA Onisoa Del Sonya",cin:"",contact:"386482664/034 30 357 00",correspondance:"",situation:"légalisé",date_signature:"2025-08-26",date_fin:"2030-08-26",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6700000.0,solde:33300000.0,mensualites_ecoulees:10.0,mensualites_restantes:50.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P31/IMT",num:85,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"31",nom:"NIRINA HARIMAMPIONONA  DANIELLA",cin:"",contact:"343740474",correspondance:"",situation:"légalisé",date_signature:"2025-08-25",date_fin:"2030-08-25",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:7370000.0,solde:32630000.0,mensualites_ecoulees:11.0,mensualites_restantes:51.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P221/IMT",num:139,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"221",nom:"RASAMBOMANANA Domoina Malala",cin:"",contact:"340567312",correspondance:"",situation:"légalisé",date_signature:"2025-10-27",date_fin:"2030-10-27",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8040000.0,solde:31960000.0,mensualites_ecoulees:12.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P84/IMT",num:142,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"84",nom:"RAHERIMAMPIONONA Noro Viviane",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6700000.0,solde:33300000.0,mensualites_ecoulees:10.0,mensualites_restantes:52.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P310/IMT",num:143,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"310",nom:"ANDRIANARIJAONA Farihy Didy",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-08-26",date_fin:"2030-08-26",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6700000.0,solde:33300000.0,mensualites_ecoulees:10.0,mensualites_restantes:52.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P146/IMT",num:145,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"146",nom:"RANDRIAMIHAJA Nambinintsoa Fanantenana",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:1340000.0,solde:38660000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P3/IMT",num:146,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"3",nom:"NIVOARILALANJATOVO",cin:"",contact:"340187930",correspondance:"",situation:"légalisé",date_signature:"2025-09-23",date_fin:"2030-09-23",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P1/IMT",num:147,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"1",nom:"RAZAFINDRABE Hanitriniana Tiana",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:2010000.0,solde:37990000.0,mensualites_ecoulees:3.0,mensualites_restantes:53.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P196/IMT",num:148,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"196",nom:"RAKOTOMALALA Heriniaina Jean Patrick",cin:"",contact:"",correspondance:"",situation:"Signé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6030000.0,solde:33970000.0,mensualites_ecoulees:9.0,mensualites_restantes:53.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P119/IMT",num:149,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"119",nom:"RAKOTOMANDIMBY  Jenny",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-07-20",date_fin:"2030-07-20",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:20160000.0,solde:19840000.0,mensualites_ecoulees:30.09,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P284/IMT",num:151,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"284",nom:"RAHERIMANANJARA Dolin Jacky",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-04",date_fin:"2030-09-04",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:6700000.0,solde:33300000.0,mensualites_ecoulees:10.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P137/IMT",num:153,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"137",nom:"RHINA Michaël",cin:"101981065893.0",contact:"331495808",correspondance:"",situation:"légalisé",date_signature:"2025-09-08",date_fin:"2030-09-08",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:8040000.0,solde:31960000.0,mensualites_ecoulees:12.0,mensualites_restantes:58.0,note:"miandry trano vita, mbola mandoa hofatrano ao @ lot 1116 E 110 Tsitamaso",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P280/IMT",num:154,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"280",nom:"RAZANAJATOVO Véronique Lovasoa",cin:"",contact:"",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:670000.0,solde:39330000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"miandry trano vita fa mandoa hofatrano sady ampiditra mpianatra au Lot 304E 284 Andafiantsimo Star",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P70/IMT",num:155,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"70",nom:"ANDRIAMAMPANDRY Jean Baptiste",cin:"103111007253.0",contact:"349534279",correspondance:"",situation:"légalisé",date_signature:"2025-09-28",date_fin:"2030-09-28",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"miandry trano vita satria mbola mandoa hofatrano au lot 304E284 Andafiantsimo Star",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P104/IMT",num:156,sigle:"IMT",localite:"Imerintsiatosika",num_lgt:"104",nom:"RAZANAMALALA Adeline",cin:"101232137998.0",contact:"343283202",correspondance:"",situation:"légalisé",date_signature:"2025-09-04",date_fin:"2030-09-04",prix_logement:40000000,mensualite:670000,nb_mensualites:60,loyer_paye:0.0,solde:40000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"mbola miandry tirage logement vao hanapakevitra hoe hanao contrat ou pas",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
];
const DATA_NSB = [
  {id:"P33/NSB",num:1,sigle:"NSB",localite:"Nosy Be",num_lgt:"33",nom:"ABDALLAH Kassim",cin:"",contact:"032 59 818 25",correspondance:"325981825.0",situation:"légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:850000,nb_mensualites:60,loyer_paye:6950000.0,solde:45050000.0,mensualites_ecoulees:8.18,mensualites_restantes:51.82,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P16/NSB",num:2,sigle:"NSB",localite:"Nosy Be",num_lgt:"16",nom:"MATHIEU Zarasoa Julia",cin:"",contact:"032 51 598 74",correspondance:"325159874.0",situation:"Signé",date_signature:"2025-05-28",date_fin:"2030-05-28",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:5220000.0,solde:46780000.0,mensualites_ecoulees:6.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P41/NSB",num:3,sigle:"NSB",localite:"Nosy Be",num_lgt:"41",nom:"RAJAOSOA Clemence",cin:"",contact:"034 41 116 12 / 032 74 554 11",correspondance:"",situation:"Signé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:1740000.0,solde:50260000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P62/NSB",num:4,sigle:"NSB",localite:"Nosy Be",num_lgt:"62",nom:"MBOTIRIZIKY Marie Witacha",cin:"",contact:"032 43 545 68 / 032 78 581 88",correspondance:"327858188.0",situation:"légalisé",date_signature:"2025-09-09",date_fin:"2030-09-09",prix_logement:52000000,mensualite:900000,nb_mensualites:60,loyer_paye:1800000.0,solde:50200000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P22/NSB",num:5,sigle:"NSB",localite:"Nosy Be",num_lgt:"22",nom:"SOATRA Maxiwell",cin:"",contact:"032 02 050 64",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P28/NSB",num:6,sigle:"NSB",localite:"Nosy Be",num_lgt:"28",nom:"JAORIZIKY Ahmad Francia",cin:"",contact:"032 77 564 27 / 034 20 685 08",correspondance:"327756427.0",situation:"légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:3480000.0,solde:48520000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P40/NSB",num:7,sigle:"NSB",localite:"Nosy Be",num_lgt:"40",nom:"MANIVA Justin Roméo",cin:"",contact:"032 63 148 64",correspondance:"",situation:"légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P8/NSB",num:9,sigle:"NSB",localite:"Nosy Be",num_lgt:"8",nom:"RAKOTOARISOA Mavana Andriamasitsaotra",cin:"",contact:"032 11 069 16",correspondance:"344706803.0",situation:"signé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:870000.0,solde:51130000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P38/NSB",num:10,sigle:"NSB",localite:"Nosy Be",num_lgt:"38",nom:"MANIFATRA Marius",cin:"",contact:"032 02 820 46",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P6/NSB",num:12,sigle:"NSB",localite:"Nosy Be",num_lgt:"6",nom:"RAMANDIMBISON Sylvestre",cin:"",contact:"034 39 452 42",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2028-05-21",prix_logement:52000000,mensualite:1450000,nb_mensualites:36,loyer_paye:5800000.0,solde:46200000.0,mensualites_ecoulees:4.0,mensualites_restantes:32.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P58/NSB",num:13,sigle:"NSB",localite:"Nosy Be",num_lgt:"58",nom:"DILIFERA Baovavy Sylvianne",cin:"",contact:"032 78 407 69",correspondance:"",situation:"légalisé",date_signature:"2025-05-22",date_fin:"2030-03-22",prix_logement:52000000,mensualite:870000,nb_mensualites:58,loyer_paye:1740000.0,solde:50260000.0,mensualites_ecoulees:2.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P68/NSB",num:14,sigle:"NSB",localite:"Nosy Be",num_lgt:"68",nom:"CHRISTELLA Assany Velonjara",cin:"",contact:"032 02 675 36",correspondance:"",situation:"légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:3480000.0,solde:48520000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P75/NSB",num:16,sigle:"NSB",localite:"Nosy Be",num_lgt:"75",nom:"RANDRIANTSALAMA Ketty Asmeralda",cin:"",contact:"032 80 045 42",correspondance:"",situation:"signé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:840000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P61/NSB",num:17,sigle:"NSB",localite:"Nosy Be",num_lgt:"61",nom:"RAKOTOARIVELO Jean Elzo",cin:"",contact:"032 02 336 35",correspondance:"",situation:"signé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P3/NSB",num:18,sigle:"NSB",localite:"Nosy Be",num_lgt:"3",nom:"LACHAUD Elgina",cin:"",contact:"032 46 441 34",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:870000.0,solde:51130000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P14/NSB",num:19,sigle:"NSB",localite:"Nosy Be",num_lgt:"14",nom:"ABDALLAH Oussouf Mananjara",cin:"",contact:"032 05 970 42",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-04-21",prix_logement:52000000,mensualite:850000,nb_mensualites:59,loyer_paye:2000000.0,solde:50000000.0,mensualites_ecoulees:2.35,mensualites_restantes:56.65,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P73/NSB",num:20,sigle:"NSB",localite:"Nosy Be",num_lgt:"73",nom:"SAMSIDIN BEFAMO Samera",cin:"",contact:"032 50 196 96 / 032 67 727 O2",correspondance:"",situation:"légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:450000,nb_mensualites:60,loyer_paye:25000000.0,solde:27000000.0,mensualites_ecoulees:55.56,mensualites_restantes:4.44,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P43/NSB",num:21,sigle:"NSB",localite:"Nosy Be",num_lgt:"43",nom:"RAZAFIHARISOA Tahiry Andrea",cin:"",contact:"032 11 761 80",correspondance:"",situation:"Signé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P24/NSB",num:22,sigle:"NSB",localite:"Nosy Be",num_lgt:"24",nom:"ANDRIANARIJAONA Andry Daniel",cin:"",contact:"032 48 023 75",correspondance:"",situation:"légalisé",date_signature:"2025-06-03",date_fin:"2030-06-03",prix_logement:52000000,mensualite:700000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P63/NSB",num:23,sigle:"NSB",localite:"Nosy Be",num_lgt:"63",nom:"TOMBO Anne Patricia cecile",cin:"",contact:"032 02 660 68",correspondance:"320266068.0",situation:"légalisé",date_signature:"2025-05-16",date_fin:"2030-05-16",prix_logement:52000000,mensualite:700000,nb_mensualites:60,loyer_paye:13500000.0,solde:38500000.0,mensualites_ecoulees:19.29,mensualites_restantes:40.71,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P76/NSB",num:24,sigle:"NSB",localite:"Nosy Be",num_lgt:"76",nom:"HICHATY",cin:"",contact:"032 57 185 61",correspondance:"325718561.0",situation:"Légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:1740000.0,solde:50260000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P52/NSB",num:25,sigle:"NSB",localite:"Nosy Be",num_lgt:"52",nom:"RAZANAJAONA Tsiry Hasina",cin:"",contact:"034 37 755 15",correspondance:"3276337019.0",situation:"légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:700000,nb_mensualites:60,loyer_paye:24000000.0,solde:28000000.0,mensualites_ecoulees:34.29,mensualites_restantes:25.71,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P54/NSB",num:26,sigle:"NSB",localite:"Nosy Be",num_lgt:"54",nom:"RABEBIARISOA Olivia Dany",cin:"",contact:"034 05 534 47 / 032 76 100 37",correspondance:"",situation:"légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P42/NSB",num:28,sigle:"NSB",localite:"Nosy Be",num_lgt:"42",nom:"BOUDI Philippine",cin:"",contact:"032 05 435  50",correspondance:"324977292.0",situation:"Signé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:1000000.0,solde:51000000.0,mensualites_ecoulees:1.15,mensualites_restantes:58.85,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P57/NSB",num:29,sigle:"NSB",localite:"Nosy Be",num_lgt:"57",nom:"RABARIVELO Mbolatiana Cael",cin:"",contact:"327679558",correspondance:"",situation:"légalisé",date_signature:"2025-05-30",date_fin:"2030-05-30",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P12/NSB",num:30,sigle:"NSB",localite:"Nosy Be",num_lgt:"12",nom:"TSARAHITA Joséphine",cin:"",contact:"032 02 081 65",correspondance:"",situation:"Légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P17/NSB",num:32,sigle:"NSB",localite:"Nosy Be",num_lgt:"17",nom:"EMERY Zafilaza",cin:"",contact:"032 06 072 92",correspondance:"320607292.0",situation:"légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:2610000.0,solde:49390000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P47/NSB",num:34,sigle:"NSB",localite:"Nosy Be",num_lgt:"47",nom:"RADY Lydie",cin:"",contact:"032 41 161 63",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:3480000.0,solde:48520000.0,mensualites_ecoulees:4.0,mensualites_restantes:56.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P2/NSB",num:35,sigle:"NSB",localite:"Nosy Be",num_lgt:"2",nom:"MANJARIA",cin:"",contact:"034 48 402 50",correspondance:"",situation:"Légalisé",date_signature:"2025-05-30",date_fin:"2030-05-30",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:870000.0,solde:51130000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P45/NSB",num:36,sigle:"NSB",localite:"Nosy Be",num_lgt:"45",nom:"VOLAHARINTIANA Germaline Eulalia",cin:"",contact:"034 00 141 65",correspondance:"",situation:"Légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P18/NSB",num:39,sigle:"NSB",localite:"Nosy Be",num_lgt:"18",nom:"RABE Jean Clauris",cin:"",contact:"032 91 764 75",correspondance:"",situation:"Légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:870000.0,solde:51130000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P72/NSB",num:40,sigle:"NSB",localite:"Nosy Be",num_lgt:"72",nom:"BEFAMO Sylvie",cin:"",contact:"032 41 711 50",correspondance:"",situation:"légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:780000,nb_mensualites:60,loyer_paye:5220000.0,solde:46780000.0,mensualites_ecoulees:6.69,mensualites_restantes:53.31,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P10/NSB",num:41,sigle:"NSB",localite:"Nosy Be",num_lgt:"10",nom:"RANDRIANASOA Feno",cin:"",contact:"034 07 115 84",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:700000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P74/NSB",num:42,sigle:"NSB",localite:"Nosy Be",num_lgt:"74",nom:"BELAZA Herve",cin:"",contact:"032 57 108 70",correspondance:"",situation:"Légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P27/NSB",num:43,sigle:"NSB",localite:"Nosy Be",num_lgt:"27",nom:"TOVOSON Andriantsiriniaina Jules",cin:"",contact:"034 38 707 14",correspondance:"",situation:"Légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P21/NSB",num:44,sigle:"NSB",localite:"Nosy Be",num_lgt:"21",nom:"JERISA William",cin:"",contact:"034 02 032 71",correspondance:"",situation:"Légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P26/NSB",num:45,sigle:"NSB",localite:"Nosy Be",num_lgt:"26",nom:"TOTO Marie Claude",cin:"",contact:"032 04 174 24",correspondance:"",situation:"légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:870000.0,solde:51130000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P20/NSB",num:48,sigle:"NSB",localite:"Nosy Be",num_lgt:"20",nom:"JULIETTE Bezandry",cin:"",contact:"032 64 449 22",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:1740000.0,solde:50260000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P15/NSB",num:51,sigle:"NSB",localite:"Nosy Be",num_lgt:"15",nom:"RANDRIANANTOANDRO Riana Naliarimanotrona",cin:"",contact:"034 12 629 08",correspondance:"",situation:"signé",date_signature:"2025-07-03",date_fin:"2030-07-03",prix_logement:52000000,mensualite:440000,nb_mensualites:60,loyer_paye:26000000.0,solde:26000000.0,mensualites_ecoulees:59.09,mensualites_restantes:0.91,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P46/NSB",num:54,sigle:"NSB",localite:"Nosy Be",num_lgt:"46",nom:"RAZANADRAVO Nathalie",cin:"",contact:"032 43 556 44",correspondance:"386522268.0",situation:"signé",date_signature:"2025-09-11",date_fin:"2030-09-11",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:1740000.0,solde:50260000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P56/NSB",num:55,sigle:"NSB",localite:"Nosy Be",num_lgt:"56",nom:"RAJESSON Hanitriniony",cin:"",contact:"034 02 017 05",correspondance:"",situation:"légalisé",date_signature:"2025-05-30",date_fin:"2030-05-30",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:5220000.0,solde:46780000.0,mensualites_ecoulees:6.0,mensualites_restantes:54.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P66/NSB",num:56,sigle:"NSB",localite:"Nosy Be",num_lgt:"66",nom:"ZANATIANA Létitia Julia",cin:"",contact:"344456169",correspondance:"344456169.0",situation:"légalisé",date_signature:"2025-05-22",date_fin:"2030-05-22",prix_logement:52000000,mensualite:540000,nb_mensualites:60,loyer_paye:20010000.0,solde:31990000.0,mensualites_ecoulees:37.06,mensualites_restantes:22.94,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P11/NSB",num:58,sigle:"NSB",localite:"Nosy Be",num_lgt:"11",nom:"RAZANAKOTO Vololonary Lalao",cin:"",contact:"032 07 789 05",correspondance:"",situation:"Signé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:2610000.0,solde:49390000.0,mensualites_ecoulees:3.0,mensualites_restantes:57.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P/NSB",num:59,sigle:"NSB",localite:"Nosy Be",num_lgt:"",nom:"RANOMENJANAHARY Nirinarimanana",cin:"",contact:"032 84 121 84",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P31/NSB",num:65,sigle:"NSB",localite:"Nosy Be",num_lgt:"31",nom:"RAKOTOVAO Félicia Romie",cin:"",contact:"034 02 301 60",correspondance:"340230160.0",situation:"Légalisé",date_signature:"2025-05-30",date_fin:"2030-05-30",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:1740000.0,solde:50260000.0,mensualites_ecoulees:2.0,mensualites_restantes:58.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P19/NSB",num:66,sigle:"NSB",localite:"Nosy Be",num_lgt:"19",nom:"SOAHARIMAMY Estelle",cin:"",contact:"032 79 812 91",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:870000.0,solde:51130000.0,mensualites_ecoulees:1.0,mensualites_restantes:59.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P80/NSB",num:67,sigle:"NSB",localite:"Nosy Be",num_lgt:"80",nom:"RAVAOARISOA Emma Fideline",cin:"",contact:"032 45 243 45",correspondance:"",situation:"Légalisé",date_signature:"2025-05-21",date_fin:"2030-05-21",prix_logement:52000000,mensualite:750000,nb_mensualites:60,loyer_paye:7000000.0,solde:45000000.0,mensualites_ecoulees:9.33,mensualites_restantes:50.67,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P1/NSB",num:69,sigle:"NSB",localite:"Nosy Be",num_lgt:"1",nom:"SINOHA Tony Olivier",cin:"",contact:"032 04 513 15",correspondance:"",situation:"Légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P25/NSB",num:70,sigle:"NSB",localite:"Nosy Be",num_lgt:"25",nom:"ZARANIRINA Nathani Casimir",cin:"",contact:"032 21 593 41 / 032 81 426 70",correspondance:"",situation:"Légalisé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
  {id:"P71/NSB",num:71,sigle:"NSB",localite:"Nosy Be",num_lgt:"71",nom:"AMIDJEE Rosemine",cin:"",contact:"032 04 787 33",correspondance:"",situation:"Signé",date_signature:"2025-05-23",date_fin:"2030-05-23",prix_logement:52000000,mensualite:870000,nb_mensualites:60,loyer_paye:0.0,solde:52000000.0,mensualites_ecoulees:0.0,mensualites_restantes:60.0,note:"",acompte:0,adresse:"",mode_quittance:"whatsapp",quittance_contact:"",region:"",type_logement:"",paiements:[]},
];
const DATA_FNR = [];
const DATA_TTV = [];

// ============================================================
// ANALOGH SVR — Suivi Recouvrement Multi-Sites
// Logo ANALOGH embarqué
// ============================================================
const LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCAKQBkADAREAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9PaACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKAKWp65ouimzXWdXsrA6jdJY2YurhIvtFy4JSGPcRvkIViFGSdpwOKaTextRw1bE8zowcuVOTsm7RW7dtkrq7ehdpGIUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAV9S1LTtH0+51bV7+2sbGziae4ubmVYooY1GWd3YgKoAJJJwKaTbsjSjRqYipGlRi5Sk7JJXbb2SS1bPjn42f8FGvDPh2a68P/BjRE8Q30LeWdZvwyaepBQkxRqRJOMGVMkxAMqsvmKee2lgnLWeh+w8OeEeJxajiM5n7OL+xHWfXd6xj0enM2m0+Vmj+xr4G8R/FnQ/EPx5+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/wAk0AH/AA8O+C3/AELHjb/wCtP/AJJoA6n4ZftlfDD4reONN8A+HdC8UW+oap53ky3trbpCvlQvK24pOzD5YyBhTyR060Ae8UAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFAHl/7T/8AyQDxx/2C3/8AQloA/KKgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgD2v9jH/k5Xwf/wBxD/0guKAP1CoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKAPL/ANp//kgHjj/sFv8A+hLQB+UVABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFAHtf7GP/ACcr4P8A+4h/6QXFAH6hUAFABQAUAFAHyN+0p+2P4y+E3xQuPAfgfSNBvIdOs7dr59TtZ2dbqRTJtQxzICnlPCemdxbmgD3D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tLc2bW+hQ3NuRLL5wUteITjanl5RTg7xK5BAA3AH2vQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAcV/wo/4Lf9Eg8E/+E/af/G6AOs03TNN0awg0rR9PtrGytUEcFtbRLFFEg6KqKAFHsBQBZoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoAKACgAoA57xF8Ovh94vvU1LxZ4E8Pa3dxRCBLjUdLguZFjBJCBpFJC5Zjjpkn1oA1NF0PRfDemQ6L4d0ey0vT7fd5NpZW6QQx7mLNtRAFGWYk4HJJPegC7QBheI/APgXxjNDceLvBWg65LbKUhfUtNhuWjUnJCmRSQCewoA0NF0PRfDemQ6L4d0ey0vT7fd5NpZW6QQx7mLNtRAFGWYk4HJJPegC7QAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQBS1rQ9F8SaZNoviLR7LVNPuNvnWl7bpPDJtYMu5HBU4ZQRkcEA9qAOdsfg78I9MvbfUtN+FnhC0u7SVJ7e4g0O1jkhkUhldGVAVYEAgjkEUAdfQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFABQAUAFAH/9k=";
const LOGO_SRC = `data:image/jpeg;base64,${LOGO_B64}`;

// Couleurs ANALOGH (bordeaux/magenta = #B5006E, gris = #555555)
const NALOGH = {
  primary:   "#B5006E",
  primaryDk: "#8a0054",
  primaryLt: "#d4007f",
  accent:    "#e8007a",
  gray:      "#555555",
  grayLt:    "#888888",
  bg:        "#0f0f1a",
  card:      "#1a1a2e",
  card2:     "#16213e",
  border:    "#2a2a4a",
  text:      "#f1f1f1",
  muted:     "#9999bb",
  dim:       "#666688",
};

// ============================================================
// AUTHENTIFICATION — Comptes utilisateurs
// ============================================================
const USERS_DEFAULT = {
  "dvr.analogh@gmail.com": {
    email:    "dvr.analogh@gmail.com",
    password: "Analogh2026!",
    nom:      "Administrateur ANALOGH",
    role:     "admin",
    actif:    true,
    },
};

const AUTH_KEY  = "svr_auth_v3";
const USERS_KEY = "svr_users_v3";
const USERS_SHARED_KEY = "svr_shared_users_v1";

// ============================================================
// CONFIGURATION DES SITES
// ============================================================
const SITES_CONFIG_DEFAULT = {
  ABE: { nom:"Antsirabe",        sigle:"ABE", localite:"Vatofotsy Antsirabe", region:"Vakinankaratra",  couleur:"#3b82f6", emoji:"🏔️", prix_defaut:40000000, mensualite_defaut:670000,  nb_logements_total:79  },
  IMT: { nom:"Imerintsiatosika", sigle:"IMT", localite:"Imerintsiatosika",    region:"Itasy",           couleur:"#10b981", emoji:"🌿", prix_defaut:40000000, mensualite_defaut:670000,  nb_logements_total:319 },
  NSB: { nom:"Nosy Be",          sigle:"NSB", localite:"Nosy Be",             region:"DIANA",           couleur:"#f59e0b", emoji:"🏝️", prix_defaut:52000000, mensualite_defaut:870000,  nb_logements_total:79  },
  FNR: { nom:"Fianarantsoa",     sigle:"FNR", localite:"Fianarantsoa",        region:"Haute Matsiatra", couleur:"#8b5cf6", emoji:"⛰️", prix_defaut:40000000, mensualite_defaut:670000,  nb_logements_total:15  },
  TTV: { nom:"Tamatave",         sigle:"TTV", localite:"Toamasina",           region:"Atsinanana",      couleur:"#ef4444", emoji:"🌊", prix_defaut:45000000, mensualite_defaut:750000,  nb_logements_total:79  },
};

const STORAGE_KEY       = "svr_2026_v5";
const STORAGE_SITES_KEY = "svr_sites_v5";
const MOIS = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
const MOIS_COURT = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"];

const fmt  = (n) => new Intl.NumberFormat('fr-MG').format(Math.round(n||0)) + " Ar";
const fmtM = (n) => { n=n||0; if(n>=1000000000) return (n/1000000000).toFixed(2)+"Md"; if(n>=1000000) return (n/1000000).toFixed(1)+"M"; if(n>=1000) return (n/1000).toFixed(0)+"k"; return String(n); };
const norm = (s) => (s||"").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
const isActif = (a) => ["legalise","signe","suspendu"].some(x => norm(a.situation).includes(x));

const INITIAL_DATA = {
  ABE: DATA_ABE.map(a=>({...a,sigle:"ABE"})),
  IMT: DATA_IMT.map(a=>({...a,sigle:"IMT"})),
  NSB: DATA_NSB.map(a=>({...a,sigle:"NSB"})),
  FNR: DATA_FNR.map(a=>({...a,sigle:"FNR"})),
  TTV: DATA_TTV.map(a=>({...a,sigle:"TTV"})),
};

// ============================================================
// APP PRINCIPALE avec authentification
// ============================================================
export default function App() {

  // ── Chargement session ──
  const [session, setSession] = useState(() => {
    try {
      const raw = localStorage.getItem(AUTH_KEY);
      if (!raw) return null;
      const s = JSON.parse(raw);
      if (!s || !s.email || !s.role || !s.nom) return null;
      return s;
    } catch { return null; }
  });

  // ── Chargement utilisateurs (partagés via données app) ──
  const [users, setUsers] = useState(() => {
    try {
      // Lire depuis la clé partagée EN PRIORITÉ (créée par l'admin)
      const shared = localStorage.getItem(USERS_SHARED_KEY);
      const local  = localStorage.getItem(USERS_KEY);
      const base   = shared ? JSON.parse(shared) : (local ? JSON.parse(local) : {});
      const adminKey = "dvr.analogh@gmail.com";
      if (!base[adminKey]) base[adminKey] = { ...USERS_DEFAULT[adminKey] };
      return base;
    } catch { return { ...USERS_DEFAULT }; }
  });

  // ── Persistance dans les DEUX clés à chaque changement ──
  useEffect(() => {
    try {
      localStorage.setItem(USERS_KEY,        JSON.stringify(users));
      localStorage.setItem(USERS_SHARED_KEY, JSON.stringify(users));
    } catch {}
  }, [users]);

  // ── Connexion ──
  function login(email, password, remember) {
    const key = email.toLowerCase().trim();
    const pwd = (password || "").trim();
    // Lire depuis TOUTES les sources pour être sûr d'avoir les derniers users
    let allUsers = { ...USERS_DEFAULT, ...users };
    try {
      const shared = localStorage.getItem(USERS_SHARED_KEY);
      const local  = localStorage.getItem(USERS_KEY);
      if (shared) allUsers = { ...allUsers, ...JSON.parse(shared) };
      else if (local) allUsers = { ...allUsers, ...JSON.parse(local) };
    } catch {}
    const u = allUsers[key];
    if (u && u.actif === true && u.password === pwd) {
      const s = { email: u.email, nom: u.nom, role: u.role };
      localStorage.setItem(AUTH_KEY, JSON.stringify(s));
      if (remember) {
        localStorage.setItem("svr_remember_v1", JSON.stringify({ email: key, password: pwd }));
      } else {
        localStorage.removeItem("svr_remember_v1");
      }
      setSession(s);
      return { ok: true };
    }
    if (!u) return { ok: false, msg: "Compte introuvable. Vérifiez l'adresse email." };
    if (!u.actif) return { ok: false, msg: "Ce compte est désactivé." };
    return { ok: false, msg: "Mot de passe incorrect." };
  }

  // ── Déconnexion ──
  function logout() {
    localStorage.removeItem(AUTH_KEY);
    setSession(null);
  }

// ── Créer un utilisateur ──
  function createUser(data) {
    const key = (data.email || "").toLowerCase().trim();
    if (!key) return { ok: false, msg: "Adresse email invalide." };
    if (!data.password || !data.password.trim()) return { ok: false, msg: "Mot de passe obligatoire." };
    if (!data.nom || !data.nom.trim()) return { ok: false, msg: "Nom obligatoire." };
    if (users[key]) return { ok: false, msg: `Le compte ${key} existe déjà.` };
    const newUser = {
      email:    key,
      password: data.password.trim(),
      nom:      data.nom.trim(),
      role:     data.role || "user",
      actif:    true,
    };
    setUsers(prev => {
      const next = { ...prev, [key]: newUser };
      // Persister immédiatement dans les deux clés
      try {
        localStorage.setItem(USERS_KEY,        JSON.stringify(next));
        localStorage.setItem(USERS_SHARED_KEY, JSON.stringify(next));
      } catch {}
      return next;
    });
    return { ok: true };
  }

  // ── Modifier un utilisateur ──
  function updateUser(email, updates) {
    const key = (email || "").toLowerCase().trim();
    if (!key) return { ok: false, msg: "Email invalide." };
    const cleanUpdates = {};
    for (const k in updates) {
      if (updates[k] !== undefined && updates[k] !== null) {
        if (k === "password" && !String(updates[k]).trim()) continue;
        cleanUpdates[k] = updates[k];
      }
    }
    setUsers(prev => {
      if (!prev[key]) return prev;
      const next = { ...prev, [key]: { ...prev[key], ...cleanUpdates } };
      try {
        localStorage.setItem(USERS_KEY,        JSON.stringify(next));
        localStorage.setItem(USERS_SHARED_KEY, JSON.stringify(next));
      } catch {}
      return next;
    });
    return { ok: true };
  }

  // ── Supprimer un utilisateur ──
  function deleteUser(email) {
    const key = (email || "").toLowerCase().trim();
    if (key === "dvr.analogh@gmail.com") return { ok: false, msg: "Impossible de supprimer l'admin principal." };
    setUsers(prev => {
      const n = { ...prev };
      delete n[key];
      try {
        localStorage.setItem(USERS_KEY,        JSON.stringify(n));
        localStorage.setItem(USERS_SHARED_KEY, JSON.stringify(n));
      } catch {}
      return n;
    });
    return { ok: true };
  }

  if (!session) return <LoginPage onLogin={login} />;

  return <MainApp session={session} onLogout={logout} users={users}
    onCreateUser={createUser} onUpdateUser={updateUser} onDeleteUser={deleteUser} />;
}



// ═══════════════════════════════════════════════════════════════════
// DESIGN SYSTEM — Inspiré iBanFirst (light, sidebar, pro)
// ═══════════════════════════════════════════════════════════════════
const DS = {
  // Fond principal blanc/gris très clair
  bg:        "#f5f6fa",
  bgPage:    "#ffffff",
  sidebar:   "#ffffff",
  card:      "#ffffff",
  cardHov:   "#f9fafb",

  // Bordures fines
  border:    "#e5e7eb",
  border2:   "#d1d5db",

  // Textes
  text:      "#111827",
  text2:     "#374151",
  text3:     "#6b7280",
  text4:     "#9ca3af",

  // Couleur marque ANALOGH
  accent:    "#B5006E",
  accentLt:  "#fdf2f8",
  accentMd:  "#fbcfe8",

  // Statuts
  green:     "#16a34a",
  greenBg:   "#f0fdf4",
  greenBd:   "#bbf7d0",
  red:       "#dc2626",
  redBg:     "#fef2f2",
  redBd:     "#fecaca",
  orange:    "#d97706",
  orangeBg:  "#fffbeb",
  orangeBd:  "#fde68a",
  blue:      "#2563eb",
  blueBg:    "#eff6ff",
  blueBd:    "#bfdbfe",

  // Sidebar
  sideW:     220,
  shadow:    "0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)",
  shadowMd:  "0 4px 6px rgba(0,0,0,0.05), 0 2px 4px rgba(0,0,0,0.04)",
  shadowLg:  "0 10px 25px rgba(0,0,0,0.08)",
};

// Formatage
const F = {
  ar:    n => { n=n||0; if(n>=1e9) return (n/1e9).toFixed(2)+" Md Ar"; if(n>=1e6) return (n/1e6).toFixed(1)+" M Ar"; if(n>=1e3) return (n/1e3).toFixed(0)+" k Ar"; return n+" Ar"; },
  full:  n => new Intl.NumberFormat('fr-MG').format(Math.round(n||0))+" Ar",
  pct:   n => (n||0).toFixed(1)+"%",
  n:     n => new Intl.NumberFormat('fr-FR').format(Math.round(n||0)),
};


// ═══════════════════════════════════════════════════════════════════
// PAGE DE LOGIN — style iBanFirst
// ═══════════════════════════════════════════════════════════════════
function LoginPage({ onLogin }) {
  const [email,    setEmail]    = useState(() => {
    try { const r = JSON.parse(localStorage.getItem("svr_remember_v1")); return r?.email||""; } catch { return ""; }
  });
  const [password, setPassword] = useState(() => {
    try { const r = JSON.parse(localStorage.getItem("svr_remember_v1")); return r?.password||""; } catch { return ""; }
  });
  const [remember, setRemember] = useState(() => !!localStorage.getItem("svr_remember_v1"));
  const [error,    setError]    = useState("");
  const [loading,  setLoading]  = useState(false);
  const [showPwd,  setShowPwd]  = useState(false);

  function handleSubmit(e) {
    e.preventDefault();
    setLoading(true); setError("");
    setTimeout(() => {
      const result = onLogin(email, password, remember);
      if (result && result.ok) {
        // ok
      } else {
        setError(result?.msg || "Identifiants incorrects.");
        setLoading(false);
      }
    }, 500);
  }

  return (
    <div style={{
      minHeight: "100vh",
      display: "grid",
      gridTemplateColumns: "1fr 1fr",
      fontFamily: "'Inter', 'Segoe UI', system-ui, sans-serif",
      background: DS.bg,
    }}>

      {/* Panneau gauche — visuel */}
      <div style={{
        background: `linear-gradient(145deg, #0f0820 0%, #1a0535 40%, #2d0b52 100%)`,
        display: "flex",
        flexDirection: "column",
        justifyContent: "space-between",
        padding: "48px",
        position: "relative",
        overflow: "hidden",
      }}>
        {/* Motif géométrique subtil */}
        <div style={{ position:"absolute", inset:0, opacity:0.04, backgroundImage:`repeating-linear-gradient(0deg, transparent, transparent 40px, rgba(255,255,255,0.5) 40px, rgba(255,255,255,0.5) 41px), repeating-linear-gradient(90deg, transparent, transparent 40px, rgba(255,255,255,0.5) 40px, rgba(255,255,255,0.5) 41px)`, pointerEvents:"none" }}/>
        {/* Cercles décoratifs */}
        <div style={{ position:"absolute", top:-120, right:-80, width:400, height:400, borderRadius:"50%", background:`radial-gradient(circle, ${DS.accent}22 0%, transparent 70%)`, pointerEvents:"none" }}/>
        <div style={{ position:"absolute", bottom:-100, left:-60, width:320, height:320, borderRadius:"50%", background:`radial-gradient(circle, #7c3aed22 0%, transparent 70%)`, pointerEvents:"none" }}/>

        {/* Logo */}
        <div style={{ position:"relative" }}>
          <img src={LOGO_SRC} alt="ANALOGH" style={{ height:44, objectFit:"contain", filter:"brightness(0) invert(1)", opacity:0.9 }}/>
        </div>

        {/* Message central */}
        <div style={{ position:"relative" }}>
          <div style={{ fontSize:11, fontWeight:600, color:DS.accent, letterSpacing:"2px", textTransform:"uppercase", marginBottom:16 }}>
            Plateforme de suivi
          </div>
          <h1 style={{ margin:0, fontSize:36, fontWeight:700, color:"#ffffff", lineHeight:1.2, letterSpacing:"-0.5px", marginBottom:16 }}>
            Suivi Recouvrement<br/>
            <span style={{ color: DS.accentMd }}>Multi-Sites</span>
          </h1>
          <p style={{ margin:0, fontSize:14, color:"rgba(255,255,255,0.55)", lineHeight:1.7, maxWidth:340 }}>
            Gérez vos dossiers acquéreurs, suivez les paiements et pilotez la performance de vos sites immobiliers en temps réel.
          </p>

          {/* Stats rapides */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginTop:40 }}>
            {[["5", "Sites"], ["161", "Dossiers"], ["100%", "Sécurisé"]].map(([v,l]) => (
              <div key={l} style={{ background:"rgba(255,255,255,0.06)", borderRadius:10, padding:"14px 16px", border:"1px solid rgba(255,255,255,0.08)" }}>
                <div style={{ fontSize:22, fontWeight:800, color:"#fff" }}>{v}</div>
                <div style={{ fontSize:10, color:"rgba(255,255,255,0.45)", marginTop:2, textTransform:"uppercase", letterSpacing:"0.8px" }}>{l}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer */}
        <div style={{ position:"relative", fontSize:11, color:"rgba(255,255,255,0.3)" }}>
          ANALOGH — Agence Nationale d'Appui au Logement et à l'Habitat · {new Date().getFullYear()}
        </div>
      </div>

      {/* Panneau droit — formulaire */}
      <div style={{ display:"flex", alignItems:"center", justifyContent:"center", padding:"48px", background:DS.bgPage }}>
        <div style={{ width:"100%", maxWidth:380 }}>

          {/* En-tête */}
          <div style={{ marginBottom:36 }}>
            <h2 style={{ margin:0, fontSize:26, fontWeight:700, color:DS.text, letterSpacing:"-0.5px" }}>
              Connexion
            </h2>
            <p style={{ margin:"8px 0 0", fontSize:14, color:DS.text3 }}>
              Accédez à votre espace de gestion
            </p>
          </div>

          <form onSubmit={handleSubmit} style={{ display:"flex", flexDirection:"column", gap:16 }}>

            {/* Email */}
            <div>
              <label style={{ display:"block", fontSize:12, fontWeight:500, color:DS.text2, marginBottom:6 }}>
                Adresse email
              </label>
              <input
                type="email"
                value={email}
                onChange={e => { setEmail(e.target.value); setError(""); }}
                placeholder="votre@email.mg"
                required
                autoFocus={!email}
                style={{
                  width:"100%", boxSizing:"border-box",
                  border:`1px solid ${error ? DS.red+"88" : DS.border2}`,
                  borderRadius:8, padding:"11px 14px",
                  fontSize:14, color:DS.text,
                  background:"#fff",
                  outline:"none", transition:"border 0.2s",
                }}
                onFocus={e => e.target.style.borderColor = DS.accent}
                onBlur={e => e.target.style.borderColor = error ? DS.red+"88" : DS.border2}
              />
            </div>

            {/* Password */}
            <div>
              <label style={{ display:"block", fontSize:12, fontWeight:500, color:DS.text2, marginBottom:6 }}>
                Mot de passe
              </label>
              <div style={{ position:"relative" }}>
                <input
                  type={showPwd ? "text" : "password"}
                  value={password}
                  onChange={e => { setPassword(e.target.value); setError(""); }}
                  placeholder="••••••••••"
                  required
                  style={{
                    width:"100%", boxSizing:"border-box",
                    border:`1px solid ${error ? DS.red+"88" : DS.border2}`,
                    borderRadius:8, padding:"11px 42px 11px 14px",
                    fontSize:14, color:DS.text,
                    background:"#fff",
                    outline:"none", transition:"border 0.2s",
                  }}
                  onFocus={e => e.target.style.borderColor = DS.accent}
                  onBlur={e => e.target.style.borderColor = error ? DS.red+"88" : DS.border2}
                />
                <button
                  type="button"
                  onClick={() => setShowPwd(!showPwd)}
                  style={{ position:"absolute", right:12, top:"50%", transform:"translateY(-50%)",
                    background:"none", border:"none", cursor:"pointer",
                    color:DS.text4, fontSize:13, padding:2 }}>
                  {showPwd ? "Masquer" : "Afficher"}
                </button>
              </div>
            </div>

            {/* Se souvenir de moi */}
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <div
                onClick={() => setRemember(!remember)}
                style={{
                  width:18, height:18, borderRadius:4, cursor:"pointer", flexShrink:0,
                  border:`2px solid ${remember ? DS.accent : DS.border2}`,
                  background: remember ? DS.accent : "#fff",
                  display:"flex", alignItems:"center", justifyContent:"center",
                  transition:"all 0.15s",
                }}>
                {remember && (
                  <svg width="10" height="8" viewBox="0 0 10 8" fill="none">
                    <path d="M1 4L3.5 6.5L9 1" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                )}
              </div>
              <span
                onClick={() => setRemember(!remember)}
                style={{ fontSize:13, color:DS.text3, cursor:"pointer", userSelect:"none" }}>
                Se souvenir de moi sur cet appareil
              </span>
            </div>

            {/* Erreur */}
            {error && (
              <div style={{
                background:DS.redBg, border:`1px solid ${DS.redBd}`,
                borderRadius:8, padding:"10px 14px",
                fontSize:13, color:DS.red,
                display:"flex", alignItems:"center", gap:8,
              }}>
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                  <circle cx="8" cy="8" r="7" stroke="#dc2626" strokeWidth="1.5"/>
                  <path d="M8 5V8.5M8 11H8.01" stroke="#dc2626" strokeWidth="1.5" strokeLinecap="round"/>
                </svg>
                {error}
              </div>
            )}

            {/* Bouton connexion */}
            <button
              type="submit"
              disabled={loading}
              style={{
                background: loading ? DS.border2 : DS.accent,
                border:"none", borderRadius:8,
                padding:"12px", color:"#fff",
                fontSize:14, fontWeight:600,
                cursor: loading ? "not-allowed" : "pointer",
                letterSpacing:"0.2px",
                transition:"all 0.2s",
                boxShadow: loading ? "none" : `0 4px 14px ${DS.accent}44`,
                marginTop:4,
              }}>
              {loading ? (
                <span style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:8 }}>
                  <svg width="16" height="16" viewBox="0 0 16 16" style={{ animation:"spin 0.8s linear infinite" }}>
                    <circle cx="8" cy="8" r="6" fill="none" stroke="rgba(255,255,255,0.3)" strokeWidth="2"/>
                    <path d="M8 2 A6 6 0 0 1 14 8" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round"/>
                  </svg>
                  Connexion en cours…
                </span>
              ) : "Se connecter"}
            </button>

          </form>

          <div style={{ marginTop:32, paddingTop:24, borderTop:`1px solid ${DS.border}`, fontSize:12, color:DS.text4, textAlign:"center", lineHeight:1.6 }}>
            Accès restreint aux utilisateurs autorisés.<br/>
            Contactez votre administrateur pour tout accès.
          </div>
        </div>
      </div>

      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        * { -webkit-font-smoothing: antialiased; }
        input, select, button { font-family: inherit; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; }
        ::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: #9ca3af; }
      `}</style>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// LAYOUT PRINCIPAL — Sidebar fixe + contenu
// ═══════════════════════════════════════════════════════════════════
function MainApp({ session, onLogout, users, onCreateUser, onUpdateUser, onDeleteUser }) {
  const [data, setData] = useState(() => {
    try {
      const s = localStorage.getItem(STORAGE_KEY);
      if (!s) return INITIAL_DATA;
      const parsed = JSON.parse(s);
      const fixed = {};
      for (const site in parsed) fixed[site] = (parsed[site]||[]).map(a=>({...a,sigle:site}));
      if (!fixed.FNR) fixed.FNR = [];
      if (!fixed.TTV) fixed.TTV = [];
      return fixed;
    } catch { return INITIAL_DATA; }
  });

  const [sitesConfig, setSitesConfig] = useState(() => {
    try { const s = localStorage.getItem(STORAGE_SITES_KEY); return s?JSON.parse(s):SITES_CONFIG_DEFAULT; }
    catch { return SITES_CONFIG_DEFAULT; }
  });

  const [vue,        setVue]        = useState("dashboard");
  const [siteActif,  setSiteActif]  = useState("tous");
  const [selectedId, setSelectedId] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [filterSit,  setFilterSit]  = useState("tous");
  const [modal,      setModal]      = useState(null);
  const [printMode,  setPrintMode]  = useState(false);
  const [moisVue,    setMoisVue]    = useState(new Date().getMonth());

  useEffect(()=>{ try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }catch{} },[data]);
  useEffect(()=>{ try{ localStorage.setItem(STORAGE_SITES_KEY, JSON.stringify(sitesConfig)); }catch{} },[sitesConfig]);

  const tous   = useMemo(()=>Object.values(data).flat(),[data]);
  const actifs = useMemo(()=>siteActif==="tous"?tous:(data[siteActif]||[]),[siteActif,tous,data]);
  const selected = tous.find(a=>a.id===selectedId);

  function calcAttendues(a) {
    if (!a.date_signature) return 0;
    const sig=new Date(a.date_signature); if(isNaN(sig)) return 0;
    const now=new Date();
    const m=(now.getFullYear()-sig.getFullYear())*12+(now.getMonth()-sig.getMonth());
    return Math.max(0,Math.min(m,a.nb_mensualites));
  }

  function getStatut(a) {
    const s=norm(a.situation);
    if (s.includes("suspendu")) return "suspendu";
    if (!["legalise","signe"].some(x=>s.includes(x))) return "inactif";
    if (a.loyer_paye===0 && calcAttendues(a)>0) return "urgente";
    const att=calcAttendues(a);
    if (att===0) return "nouveau";
    if (a.mensualites_ecoulees<att-1) return "retard";
    if (a.mensualites_ecoulees<att)   return "attention";
    return "regulier";
  }

  const STATUT = {
    regulier: {bg:DS.greenBg,bd:DS.greenBd,text:DS.green,   dot:DS.green,   label:"Régulier"},
    attention:{bg:DS.orangeBg,bd:DS.orangeBd,text:DS.orange, dot:DS.orange,  label:"À surveiller"},
    retard:   {bg:DS.redBg,  bd:DS.redBd,  text:DS.red,    dot:DS.red,     label:"En retard"},
    urgente:  {bg:"#fff1f2", bd:"#fecdd3", text:"#be123c",  dot:"#be123c",  label:"Urgent"},
    nouveau:  {bg:DS.blueBg, bd:DS.blueBd, text:DS.blue,   dot:DS.blue,    label:"Nouveau"},
    suspendu: {bg:DS.orangeBg,bd:DS.orangeBd,text:DS.orange,dot:DS.orange, label:"Suspendu"},
    inactif:  {bg:"#f9fafb", bd:DS.border, text:DS.text4,  dot:DS.text4,   label:"Inactif"},
  };

  function savePaiement(acquId, p) {
    setData(prev=>{ const n={}; for(const s in prev) n[s]=prev[s].map(a=>{ if(a.id!==acquId) return a; const pm=[...(a.paiements||[]),{...p,id:Date.now()}]; const nl=a.loyer_paye+p.montant,nm=a.mensualite>0?Math.floor(nl/a.mensualite):a.mensualites_ecoulees; return{...a,paiements:pm,loyer_paye:nl,solde:a.prix_logement-nl,mensualites_ecoulees:nm,mensualites_restantes:a.nb_mensualites-nm}; }); return n; });
    setModal(null);
  }

  function deletePaiement(acquId, pId) {
    setData(prev=>{ const n={}; for(const s in prev) n[s]=prev[s].map(a=>{ if(a.id!==acquId) return a; const p=(a.paiements||[]).find(x=>x.id===pId); if(!p) return a; const pm=(a.paiements||[]).filter(x=>x.id!==pId); const nl=a.loyer_paye-p.montant,nm=a.mensualite>0?Math.floor(nl/a.mensualite):a.mensualites_ecoulees; return{...a,paiements:pm,loyer_paye:nl,solde:a.prix_logement-nl,mensualites_ecoulees:nm,mensualites_restantes:a.nb_mensualites-nm}; }); return n; });
  }

  function updateAcquereur(acquId, updates, oldSigle) {
    setData(prev=>{ const n={}; for(const s in prev) n[s]=[...prev[s]]; const os=oldSigle||Object.keys(prev).find(s=>prev[s].some(a=>a.id===acquId)); if(!os) return prev; const ao=prev[os].find(a=>a.id===acquId); if(!ao) return prev; const ns=updates.sigle||os; const upd={...ao,...updates,sigle:ns}; if(updates.prix_logement!==undefined) upd.solde=Number(updates.prix_logement)-ao.loyer_paye; if(updates.date_signature||updates.nb_mensualites){const sig=updates.date_signature||ao.date_signature,nb=updates.nb_mensualites||ao.nb_mensualites;if(sig){const d=new Date(sig);d.setMonth(d.getMonth()+parseInt(nb));upd.date_fin=d.toISOString().split("T")[0];}} if(ns!==os){n[os]=(prev[os]||[]).filter(a=>a.id!==acquId);n[ns]=[...(prev[ns]||[]),upd];}else{n[os]=(prev[os]||[]).map(a=>a.id===acquId?upd:a);} return n; });
    setModal(null);
  }

  function addSite(cfg){ setSitesConfig(prev=>({...prev,[cfg.sigle]:cfg})); setData(prev=>({...prev,[cfg.sigle]:[]})); setModal(null); }

  function addAcquereur(acq) {
    setData(prev=>{ const site=acq.sigle,liste=prev[site]||[]; const ne={...acq,sigle:site,id:acq.id||`P${acq.num_lgt}/${site}`,num:liste.length+1,loyer_paye:parseFloat(acq.acompte)||0,solde:(parseFloat(acq.prix_logement)||0)-(parseFloat(acq.acompte)||0),mensualites_ecoulees:0,mensualites_restantes:parseInt(acq.nb_mensualites)||60,paiements:[]}; return{...prev,[site]:[...liste,ne]}; });
    setModal(null);
  }

  function exporterExcel() {
    try {
      const wb=XLSX.utils.book_new(),now=new Date(),annee=now.getFullYear();
      const SO=["IMERINTSIATOSIKA","ANTSIRABE","NOSY BE","TAMATAVE","FIANARANTSOA"];
      const SM={"IMERINTSIATOSIKA":"IMT","ANTSIRABE":"ABE","NOSY BE":"NSB","TAMATAVE":"TTV","FIANARANTSOA":"FNR"};
      const pmtS={};
      SO.forEach(sn=>{ const sg=SM[sn]; pmtS[sg]=Array(12).fill(0); (data[sg]||[]).forEach(a=>(a.paiements||[]).forEach(p=>{ if(!p.date) return; const d=new Date(p.date); if(d.getFullYear()===annee) pmtS[sg][d.getMonth()]+=p.montant||0; })); });
      const tR=tous.reduce((s,a)=>s+a.loyer_paye,0),tA=tous.reduce((s,a)=>s+a.mensualite*calcAttendues(a),0);
      const SR=SO.map(sn=>{ const sg=SM[sn],cfg=sitesConfig[sg]||{},lst=data[sg]||[],rec=lst.reduce((s,a)=>s+a.loyer_paye,0),att=lst.reduce((s,a)=>s+a.mensualite*calcAttendues(a),0),prix=lst.reduce((s,a)=>s+a.prix_logement,0),nb=cfg.nb_logements_total||lst.length; return[sn,nb,lst.length,nb-lst.length,prix,att,rec,Math.max(0,att-rec),Math.max(0,lst.reduce((s,a)=>s+a.solde,0)-(att-rec))]; });
      const rows=[
        [`SITUATION GLOBALE AU ${now.toLocaleDateString("fr-FR")}`,...Array(8).fill("")],
        ["SITE","Logement Total","Logement Vendu","Logement disponible","Prix Total (Ar)","Paiement Attendu (Ar)","Paiement reçu (Ar)","Paiement non respectée (Ar)","Paiement non échu (Ar)"],
        ...SR,
        ["TOTAL",SO.reduce((s,sn)=>(sitesConfig[SM[sn]]?.nb_logements_total||0)+s,0),tous.length,"",tous.reduce((s,a)=>s+a.prix_logement,0),tA,tR,Math.max(0,tA-tR),""],
        [],["PERFORMANCE PAR SITE — Paiements reçus par mois (Ar)",...Array(8).fill("")],
        ["Site",...SO,"TOTAL"],
        ...MOIS.map((m,i)=>[m,...SO.map(sn=>pmtS[SM[sn]][i]||0),SO.reduce((s,sn)=>(pmtS[SM[sn]][i]||0)+s,0)]),
        ["TOTAL",...SO.map(sn=>pmtS[SM[sn]].reduce((s,x)=>s+x,0)),SO.reduce((s,sn)=>pmtS[SM[sn]].reduce((ss,x)=>ss+x,0)+s,0)],
        [],["TAUX DE RECOUVREMENT",...Array(8).fill("")],
        ["Site",...SO,"Moyenne"],
        ...MOIS.map((m,i)=>{ const v=SO.map(sn=>{ const a=(data[SM[sn]]||[]).reduce((s,a)=>s+a.mensualite,0); return a>0?(pmtS[SM[sn]][i]||0)/a:0; }); return[m,...v,v.filter(x=>x>0).reduce((s,x)=>s+x,0)/(v.filter(x=>x>0).length||1)]; }),
      ];
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:24},{wch:14},{wch:14},{wch:16},{wch:18},{wch:18},{wch:16},{wch:22},{wch:16}];
      XLSX.utils.book_append_sheet(wb,ws,"TDB");
      const CH=["Réf. Contrat","N° Logement","Nom Complet","Contact","Situation","Date Signature","Date Fin","Prix Logement","Mensualité","Solde initial 2026","Janvier","Fevrier","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre","Paiement 2026","Solde courant"];
      SO.forEach(sn=>{ const sg=SM[sn],lst=data[sg]||[]; if(!lst.length) return; const dr=lst.map(a=>{ const pm=Array(12).fill(0); (a.paiements||[]).forEach(p=>{ if(!p.date) return; const d=new Date(p.date); if(d.getFullYear()===annee) pm[d.getMonth()]+=p.montant||0; }); const pt=pm.reduce((s,x)=>s+x,0); return[a.id||"",a.num_lgt||"",a.nom||"",a.contact||"",a.situation||"",a.date_signature||"",a.date_fin||"",a.prix_logement||0,a.mensualite||0,(a.solde||0)+pt,...pm,pt,a.solde||0]; }); const ws2=XLSX.utils.aoa_to_sheet([CH,...dr]); ws2["!cols"]=[{wch:16},{wch:10},{wch:30},{wch:16},{wch:12},{wch:14},{wch:14},{wch:16},{wch:14},{wch:16},...Array(12).fill({wch:12}),{wch:14},{wch:14}]; XLSX.utils.book_append_sheet(wb,ws2,sn); });
      const buf=XLSX.write(wb,{bookType:"xlsx",type:"array"});
      saveAs(new Blob([buf],{type:"application/octet-stream"}),`ANALOGH_SVR_Export_${now.toISOString().split("T")[0]}.xlsx`);
    } catch(err) { console.error(err); alert("Erreur export : "+err.message); }
  }

  const alertes=useMemo(()=>actifs.filter(a=>["urgente","retard","attention"].includes(getStatut(a))).sort((a,b)=>({urgente:0,retard:1,attention:2})[getStatut(a)]-({urgente:0,retard:1,attention:2})[getStatut(b)]),[actifs]);
  const filteredList=useMemo(()=>actifs.filter(a=>{ const ms=(searchTerm||"").toLowerCase(); return(!ms||a.nom.toLowerCase().includes(ms)||a.num_lgt.includes(ms)||(a.contact||"").includes(ms))&&(filterSit==="tous"||norm(a.situation).includes(norm(filterSit))); }),[actifs,searchTerm,filterSit]);

  if (printMode && selected) return <FicheImprimable a={selected} onClose={()=>setPrintMode(false)} fmt={F.full} calcAttendues={calcAttendues} getStatut={getStatut} STATUT={STATUT} sitesConfig={sitesConfig}/>;

  const NAV = [
    { key:"dashboard", icon:<IconDashboard/>, label:"Tableau de bord" },
    { key:"liste",     icon:<IconList/>,      label:"Acquéreurs" },
    { key:"saisie",    icon:<IconEdit/>,      label:"Saisie paiement" },
    { key:"alertes",   icon:<IconBell/>,      label:"Alertes", badge:alertes.length },
    ...(session.role==="admin"?[{ key:"admin", icon:<IconSettings/>, label:"Administration" }]:[]),
  ];

  return (
    <div style={{ display:"flex", minHeight:"100vh", fontFamily:"'Inter','Segoe UI',system-ui,sans-serif", background:DS.bg }}>
      <style>{`
        * { -webkit-font-smoothing: antialiased; box-sizing: border-box; }
        input, select, button, textarea { font-family: inherit; }
        ::-webkit-scrollbar { width:5px; height:5px; }
        ::-webkit-scrollbar-track { background:transparent; }
        ::-webkit-scrollbar-thumb { background:#d1d5db; border-radius:3px; }
        @keyframes fadeIn { from{opacity:0;transform:translateY(4px)} to{opacity:1;transform:translateY(0)} }
        .page-in { animation: fadeIn 0.2s ease forwards; }
      `}</style>

      {/* ── SIDEBAR ── */}
      <aside style={{
        width: DS.sideW,
        background: DS.sidebar,
        borderRight: `1px solid ${DS.border}`,
        display: "flex",
        flexDirection: "column",
        position: "fixed",
        top: 0, left: 0, bottom: 0,
        zIndex: 50,
        boxShadow: DS.shadow,
      }}>
        {/* Logo */}
        <div style={{ padding:"20px 20px 16px", borderBottom:`1px solid ${DS.border}` }}>
          <img src={LOGO_SRC} alt="ANALOGH" style={{ height:34, objectFit:"contain" }}/>
          <div style={{ fontSize:10, color:DS.text4, marginTop:4, letterSpacing:"0.5px" }}>
            SVR · Suivi Recouvrement
          </div>
        </div>

        {/* Filtre site */}
        <div style={{ padding:"12px 14px", borderBottom:`1px solid ${DS.border}` }}>
          <div style={{ fontSize:9, fontWeight:600, color:DS.text4, textTransform:"uppercase", letterSpacing:"0.8px", marginBottom:7 }}>Site actif</div>
          <select value={siteActif} onChange={e=>setSiteActif(e.target.value)} style={{
            width:"100%", background:"#f9fafb", border:`1px solid ${DS.border}`,
            borderRadius:6, padding:"6px 9px", fontSize:12, color:DS.text2, cursor:"pointer", outline:"none",
          }}>
            <option value="tous">Tous les sites ({tous.length})</option>
            {Object.entries(sitesConfig).map(([sigle,cfg])=>(
              <option key={sigle} value={sigle}>{cfg.nom} ({(data[sigle]||[]).length})</option>
            ))}
          </select>
        </div>

        {/* Navigation */}
        <nav style={{ flex:1, padding:"10px 8px", overflowY:"auto" }}>
          {NAV.map(({ key, icon, label, badge }) => {
            const active = vue === key;
            return (
              <button key={key} onClick={()=>setVue(key)} style={{
                width:"100%", display:"flex", alignItems:"center", gap:10,
                padding:"9px 12px", borderRadius:7, marginBottom:2,
                background: active ? DS.accentLt : "transparent",
                border:"none",
                color: active ? DS.accent : DS.text3,
                fontSize:13, fontWeight: active ? 600 : 400,
                cursor:"pointer", textAlign:"left",
                transition:"all 0.15s",
              }}
              onMouseEnter={e=>{ if(!active) e.currentTarget.style.background="#f9fafb"; }}
              onMouseLeave={e=>{ if(!active) e.currentTarget.style.background="transparent"; }}>
                <span style={{ color: active ? DS.accent : DS.text4, flexShrink:0 }}>{icon}</span>
                <span style={{ flex:1 }}>{label}</span>
                {badge > 0 && (
                  <span style={{ background:DS.red, color:"#fff", fontSize:10, fontWeight:700,
                    borderRadius:999, minWidth:18, height:18, display:"flex", alignItems:"center",
                    justifyContent:"center", padding:"0 5px" }}>
                    {badge}
                  </span>
                )}
              </button>
            );
          })}

          <div style={{ height:1, background:DS.border, margin:"10px 4px" }}/>

          {/* Actions rapides */}
          <div style={{ fontSize:9, fontWeight:600, color:DS.text4, textTransform:"uppercase", letterSpacing:"0.8px", padding:"4px 12px 6px" }}>Actions</div>
          {[
            { label:"+ Acquéreur", onClick:()=>setModal({type:"newAcquereur"}), color:DS.accent },
            { label:"+ Nouveau site", onClick:()=>setModal({type:"newSite"}), color:DS.text3 },
            { label:"Export Excel", onClick:exporterExcel, color:DS.green },
          ].map(({label,onClick,color})=>(
            <button key={label} onClick={onClick} style={{
              width:"100%", padding:"8px 12px", borderRadius:7, marginBottom:2,
              background:"transparent", border:"none",
              color, fontSize:12, fontWeight:500, cursor:"pointer", textAlign:"left",
              display:"flex", alignItems:"center", gap:8,
            }}
            onMouseEnter={e=>e.currentTarget.style.background="#f9fafb"}
            onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              {label}
            </button>
          ))}
        </nav>

        {/* Profil utilisateur */}
        <div style={{ padding:"12px 14px", borderTop:`1px solid ${DS.border}` }}>
          <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:8 }}>
            <div style={{
              width:32, height:32, borderRadius:"50%",
              background: session.role==="admin" ? DS.accent : DS.blue,
              display:"flex", alignItems:"center", justifyContent:"center",
              color:"#fff", fontSize:12, fontWeight:700, flexShrink:0,
            }}>
              {(session.nom||"U").charAt(0).toUpperCase()}
            </div>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ fontSize:12, fontWeight:600, color:DS.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                {session.nom}
              </div>
              <div style={{ fontSize:10, color:DS.text4 }}>
                {session.role==="admin" ? "Administrateur" : "Utilisateur"}
              </div>
            </div>
          </div>
          <button onClick={onLogout} style={{
            width:"100%", padding:"6px 10px", borderRadius:6,
            background:"#f9fafb", border:`1px solid ${DS.border}`,
            color:DS.text3, fontSize:11, cursor:"pointer", textAlign:"center",
          }}
          onMouseEnter={e=>{e.currentTarget.style.background=DS.redBg;e.currentTarget.style.color=DS.red;e.currentTarget.style.borderColor=DS.redBd;}}
          onMouseLeave={e=>{e.currentTarget.style.background="#f9fafb";e.currentTarget.style.color=DS.text3;e.currentTarget.style.borderColor=DS.border;}}>
            Se déconnecter
          </button>
        </div>
      </aside>

      {/* ── CONTENU PRINCIPAL ── */}
      <div style={{ marginLeft:DS.sideW, flex:1, minWidth:0 }}>
        {/* Topbar */}
        <div style={{
          background:DS.bgPage, borderBottom:`1px solid ${DS.border}`,
          padding:"0 28px", height:56,
          display:"flex", alignItems:"center", justifyContent:"space-between",
          position:"sticky", top:0, zIndex:40,
        }}>
          <div>
            <div style={{ fontSize:15, fontWeight:600, color:DS.text }}>
              {NAV.find(n=>n.key===vue)?.label || ""}
            </div>
            {siteActif!=="tous" && (
              <div style={{ fontSize:11, color:DS.text4 }}>
                {sitesConfig[siteActif]?.nom} · {actifs.length} dossier{actifs.length>1?"s":""}
              </div>
            )}
          </div>
          <div style={{ fontSize:12, color:DS.text4 }}>
            {new Date().toLocaleDateString("fr-FR",{day:"numeric",month:"long",year:"numeric"})}
          </div>
        </div>

        {/* Page */}
        <main style={{ padding:24, maxWidth:1400, margin:"0 auto" }} className="page-in">
          {vue==="dashboard" && <Dashboard actifs={actifs} data={data} sitesConfig={sitesConfig} siteActif={siteActif} alertes={alertes} getStatut={getStatut} STATUT={STATUT} calcAttendues={calcAttendues} moisVue={moisVue} setMoisVue={setMoisVue} onFiche={(id)=>{setSelectedId(id);setVue("fiche");}} tous={tous}/>}
          {vue==="liste" && <ListeView actifs={actifs} filteredList={filteredList} searchTerm={searchTerm} setSearchTerm={setSearchTerm} filterSit={filterSit} setFilterSit={setFilterSit} siteActif={siteActif} sitesConfig={sitesConfig} getStatut={getStatut} STATUT={STATUT} onFiche={(id)=>{setSelectedId(id);setVue("fiche");}} onPmt={(id)=>{setSelectedId(id);setModal({type:"paiement"});}} onEdit={(a)=>setModal({type:"editAcquereur",data:a})}/>}
          {vue==="saisie" && <SaisieView acquéreurs={actifs} onSave={savePaiement} onFiche={(id)=>{setSelectedId(id);setVue("fiche");}} onEdit={(a)=>setModal({type:"editAcquereur",data:a})} getStatut={getStatut} STATUT={STATUT} sitesConfig={sitesConfig}/>}
          {vue==="alertes" && <AlertesView alertes={alertes} getStatut={getStatut} STATUT={STATUT} sitesConfig={sitesConfig} calcAttendues={calcAttendues} onPmt={(id)=>{setSelectedId(id);setModal({type:"paiement"});}} onFiche={(id)=>{setSelectedId(id);setVue("fiche");}} onEdit={(a)=>setModal({type:"editAcquereur",data:a})}/>}
          {vue==="fiche" && selected && <FicheDetaillee a={selected} onPaiement={()=>setModal({type:"paiement"})} onPrint={()=>setPrintMode(true)} onDelete={deletePaiement} onEdit={()=>setModal({type:"editAcquereur",data:selected})} getStatut={getStatut} STATUT={STATUT} calcAttendues={calcAttendues} onBack={()=>setVue("liste")} sitesConfig={sitesConfig}/>}
          {vue==="admin" && session.role==="admin" && <AdminView users={users} session={session} onCreateUser={onCreateUser} onUpdateUser={onUpdateUser} onDeleteUser={onDeleteUser}/>}
        </main>
      </div>

      {modal?.type==="paiement"      && selected    && <ModalPaiement a={selected} onSave={savePaiement} onClose={()=>setModal(null)}/>}
      {modal?.type==="newSite"                      && <ModalNouveauSite onSave={addSite} onClose={()=>setModal(null)}/>}
      {modal?.type==="newAcquereur"                 && <ModalAcquereur sitesConfig={sitesConfig} onSave={addAcquereur} onClose={()=>setModal(null)} mode="new"/>}
      {modal?.type==="editAcquereur" && modal.data  && <ModalAcquereur sitesConfig={sitesConfig} onSave={(u)=>updateAcquereur(modal.data.id,u,modal.data.sigle)} onClose={()=>setModal(null)} mode="edit" initial={modal.data}/>}
    </div>
  );
}

// ── Icônes SVG minimalistes ──
const IconDashboard = () => <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="1" y="1" width="6" height="6" rx="1" stroke="currentColor" strokeWidth="1.5"/><rect x="9" y="1" width="6" height="6" rx="1" stroke="currentColor" strokeWidth="1.5"/><rect x="1" y="9" width="6" height="6" rx="1" stroke="currentColor" strokeWidth="1.5"/><rect x="9" y="9" width="6" height="6" rx="1" stroke="currentColor" strokeWidth="1.5"/></svg>;
const IconList = () => <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><line x1="1" y1="4" x2="15" y2="4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/><line x1="1" y1="8" x2="15" y2="8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/><line x1="1" y1="12" x2="9" y2="12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>;
const IconEdit = () => <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M11 2L14 5L5 14H2V11L11 2Z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round"/></svg>;
const IconBell = () => <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 1.5C5.5 1.5 3.5 3.5 3.5 6V10L2 12h12l-1.5-2V6c0-2.5-2-4.5-4.5-4.5z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round"/><path d="M6.5 12.5a1.5 1.5 0 003 0" stroke="currentColor" strokeWidth="1.5"/></svg>;
const IconSettings = () => <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="2.5" stroke="currentColor" strokeWidth="1.5"/><path d="M8 1v2M8 13v2M1 8h2M13 8h2M3.1 3.1l1.4 1.4M11.5 11.5l1.4 1.4M11.5 4.5l1.4-1.4M3.1 12.9l1.4-1.4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>;



// ═══════════════════════════════════════════════════════════════════
// GRAPHIQUES SVG NATIFS — Thème light
// ═══════════════════════════════════════════════════════════════════
function BarChart({ data, height=130 }) {
  const max = Math.max(...data.map(d=>d.value), 1);
  const W = 36, gap = 4;
  const total = data.length;
  return (
    <svg viewBox={`0 0 ${total*(W+gap)} ${height+22}`} style={{width:"100%",height:height+22,overflow:"visible"}}>
      {data.map((d,i) => {
        const bh = Math.max(2,(d.value/max)*height);
        const x = i*(W+gap);
        const y = height-bh;
        return (
          <g key={i}>
            <rect x={x} y={y} width={W} height={bh} rx={3}
              fill={d.highlight ? DS.accent : d.value>0 ? "#fce7f3" : "#f3f4f6"}/>
            {d.highlight && <rect x={x} y={y} width={W} height={4} rx={2} fill={DS.accent}/>}
            <text x={x+W/2} y={height+14} textAnchor="middle" fontSize={9} fill={DS.text4} fontFamily="system-ui">{d.label}</text>
            {d.value > 0 && (
              <text x={x+W/2} y={y-4} textAnchor="middle" fontSize={8} fill={d.highlight?DS.accent:DS.text4} fontFamily="system-ui">
                {F.ar(d.value).replace(" Ar","")}
              </text>
            )}
          </g>
        );
      })}
    </svg>
  );
}

function MiniLineChart({ data, height=50, color=DS.accent }) {
  if (!data || data.length < 2) return null;
  const max = Math.max(...data.map(d=>d.value),1);
  const W=400, H=height;
  const pts = data.map((d,i)=>({ x:(i/(data.length-1))*W, y:H-(d.value/max)*H*0.85-2 }));
  const path = pts.map((p,i)=>`${i===0?"M":"L"}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(" ");
  const fillPath = `${path} L${pts[pts.length-1].x},${H} L0,${H} Z`;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{width:"100%",height,display:"block"}} preserveAspectRatio="none">
      <defs>
        <linearGradient id={`lg${color.replace(/[^a-z0-9]/gi,"")}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity="0.15"/>
          <stop offset="100%" stopColor={color} stopOpacity="0"/>
        </linearGradient>
      </defs>
      <path d={fillPath} fill={`url(#lg${color.replace(/[^a-z0-9]/gi,"")})`}/>
      <path d={path} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  );
}

function DonutChart({ segments, size=80 }) {
  const total = segments.reduce((s,x)=>s+x.value,0)||1;
  let angle = -90;
  const cx=size/2, cy=size/2, r=size*0.4, ri=size*0.27;
  const arcs = segments.map(seg => {
    const pct=seg.value/total, a1=angle;
    angle += pct*360;
    const toR=a=>(a*Math.PI)/180;
    const x1=cx+r*Math.cos(toR(a1)),y1=cy+r*Math.sin(toR(a1));
    const x2=cx+r*Math.cos(toR(angle)),y2=cy+r*Math.sin(toR(angle));
    const xi1=cx+ri*Math.cos(toR(a1)),yi1=cy+ri*Math.sin(toR(a1));
    const xi2=cx+ri*Math.cos(toR(angle)),yi2=cy+ri*Math.sin(toR(angle));
    const large=pct>0.5?1:0;
    const d=pct<0.01?"":
      `M${x1.toFixed(2)},${y1.toFixed(2)} A${r},${r} 0 ${large},1 ${x2.toFixed(2)},${y2.toFixed(2)} L${xi2.toFixed(2)},${yi2.toFixed(2)} A${ri},${ri} 0 ${large},0 ${xi1.toFixed(2)},${yi1.toFixed(2)} Z`;
    return {...seg,d};
  });
  return (
    <svg viewBox={`0 0 ${size} ${size}`} style={{width:size,height:size,flexShrink:0}}>
      {arcs.map((a,i)=>a.d&&<path key={i} d={a.d} fill={a.color} stroke="#fff" strokeWidth={1}/>)}
    </svg>
  );
}

function ProgressBar({ pct, color=DS.accent, height=6 }) {
  const safe = Math.min(100, Math.max(0, pct||0));
  return (
    <div style={{background:"#f3f4f6",borderRadius:999,height,overflow:"hidden"}}>
      <div style={{width:`${safe}%`,height:"100%",background:color,borderRadius:999,transition:"width 0.5s ease"}}/>
    </div>
  );
}

// ─── Card ───
function Card({ children, style, padding="18px 20px", onClick }) {
  return (
    <div
      onClick={onClick}
      style={{ background:DS.card, borderRadius:10, border:`1px solid ${DS.border}`,
        boxShadow:DS.shadow, padding, ...style }}>
      {children}
    </div>
  );
}

// ─── Badge statut ───
function StatutBadge({ statut, STATUT }) {
  const s = STATUT[statut] || STATUT.inactif;
  return (
    <span style={{ background:s.bg, color:s.text, border:`1px solid ${s.bd}`,
      fontSize:10, fontWeight:600, padding:"2px 8px", borderRadius:4, whiteSpace:"nowrap" }}>
      {s.label}
    </span>
  );
}

// ─── Bouton ───
function Btn({ children, onClick, variant="ghost", color=DS.accent, disabled=false, size="md", style:xtra={} }) {
  const sz = { sm:{padding:"4px 10px",fontSize:11}, md:{padding:"7px 14px",fontSize:12}, lg:{padding:"10px 20px",fontSize:13} }[size]||{};
  const variants = {
    ghost:   { background:"#f9fafb", border:`1px solid ${DS.border}`, color:DS.text3 },
    outline: { background:color+"0d", border:`1px solid ${color}55`, color },
    solid:   { background:color, border:"none", color:"#fff", boxShadow:`0 1px 3px ${color}44` },
    danger:  { background:"#fff", border:`1px solid ${DS.redBd}`, color:DS.red },
  };
  const s = variants[variant]||variants.ghost;
  return (
    <button onClick={onClick} disabled={disabled} style={{
      ...s,...sz, borderRadius:6, fontWeight:500, cursor:disabled?"not-allowed":"pointer",
      opacity:disabled?0.5:1, transition:"all 0.15s", outline:"none", whiteSpace:"nowrap", ...xtra
    }}
    onMouseEnter={e=>{ if(!disabled) e.currentTarget.style.opacity="0.8"; }}
    onMouseLeave={e=>{ e.currentTarget.style.opacity="1"; }}>
      {children}
    </button>
  );
}

// ─── Input ───
function Input({ value, onChange, placeholder, type="text", disabled, style:xtra={} }) {
  return (
    <input type={type} value={value} onChange={onChange} placeholder={placeholder} disabled={disabled}
      style={{ width:"100%", boxSizing:"border-box",
        background:disabled?"#f9fafb":"#fff",
        border:`1px solid ${DS.border2}`, borderRadius:6,
        padding:"8px 11px", fontSize:13, color:DS.text, outline:"none",
        opacity:disabled?0.6:1, cursor:disabled?"not-allowed":"text", ...xtra }}
      onFocus={e=>{ if(!disabled) e.target.style.borderColor=DS.accent; }}
     
    />
  );
}

// ── Séparateur section ──
function SectionLabel({ children }) {
  return <div style={{ fontSize:10,fontWeight:600,color:DS.text4,textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:10 }}>{children}</div>;
}

// ═══════════════════════════════════════════════════════════════════
// TABLEAU DE BORD
// ═══════════════════════════════════════════════════════════════════
function Dashboard({ actifs, data, sitesConfig, siteActif, alertes, getStatut, STATUT, calcAttendues, moisVue, setMoisVue, onFiche, tous }) {
  const now = new Date(), annee = now.getFullYear();

  const stats = useMemo(() => {
    const rec = actifs.reduce((s,a)=>s+a.loyer_paye,0);
    const sol = actifs.reduce((s,a)=>s+a.solde,0);
    const att = actifs.reduce((s,a)=>s+a.mensualite*calcAttendues(a),0);
    const ret = actifs.filter(a=>["urgente","retard"].includes(getStatut(a))).length;
    const reg = actifs.filter(a=>getStatut(a)==="regulier").length;
    const taux = att>0?(rec/att)*100:0;
    return { total:actifs.length, rec, sol, att, ret, reg, taux };
  },[actifs]);

  const pmtMois = useMemo(()=>MOIS.map((_,i)=>actifs.flatMap(a=>(a.paiements||[]).filter(p=>{const d=new Date(p.date||"");return d.getMonth()===i&&d.getFullYear()===annee;})).reduce((s,p)=>s+p.montant,0)),[actifs]);

  const pmtSem = useMemo(()=>[{d:1,f:7},{d:8,f:14},{d:15,f:21},{d:22,f:31}].map((sem,si)=>{
    const pmts=actifs.flatMap(a=>(a.paiements||[]).filter(p=>{const d=new Date(p.date||"");return d.getMonth()===moisVue&&d.getFullYear()===annee&&d.getDate()>=sem.d&&d.getDate()<=sem.f;}));
    return{label:`S${si+1}`,value:pmts.reduce((s,p)=>s+p.montant,0),nb:pmts.length};
  }),[actifs,moisVue]);

  const perfSites = useMemo(()=>Object.entries(data).map(([sigle,liste])=>{
    const cfg=sitesConfig[sigle]||{};
    const rec=liste.reduce((s,a)=>s+a.loyer_paye,0);
    const att=liste.reduce((s,a)=>s+a.mensualite*calcAttendues(a),0);
    const sol=liste.reduce((s,a)=>s+a.solde,0);
    const al=liste.filter(a=>["urgente","retard"].includes(getStatut(a))).length;
    return{sigle,cfg,vendu:liste.length,total:cfg.nb_logements_total||liste.length,rec,att,sol,taux:att>0?(rec/att)*100:0,al};
  }).filter(s=>s.total>0),[data,sitesConfig]);

  const repStatut = useMemo(()=>{const m={};actifs.forEach(a=>{const s=getStatut(a);m[s]=(m[s]||0)+1;});return m;},[actifs]);
  const donutSegs = Object.entries(repStatut).filter(([k])=>k!=="inactif"&&repStatut[k]>0).map(([k,v])=>({value:v,color:STATUT[k]?.dot||DS.text4}));

  const STATUT_ORD = {urgente:0,retard:1,attention:2,nouveau:3,regulier:4,suspendu:5,inactif:6};

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>

      {/* ── En-tête ── */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <h1 style={{margin:0,fontSize:18,fontWeight:700,color:DS.text,letterSpacing:"-0.3px"}}>
            {siteActif==="tous" ? "Vue d'ensemble" : sitesConfig[siteActif]?.nom}
          </h1>
          <div style={{fontSize:12,color:DS.text3,marginTop:3}}>
            {now.toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})} · Exercice {annee}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:DS.text3}}>Période :</span>
          <select value={moisVue} onChange={e=>setMoisVue(+e.target.value)} style={{
            background:"#fff",border:`1px solid ${DS.border2}`,borderRadius:6,
            padding:"6px 10px",fontSize:12,color:DS.text2,cursor:"pointer",outline:"none",
          }}>
            {MOIS.map((m,i)=><option key={i} value={i}>{m} {annee}</option>)}
          </select>
        </div>
      </div>

      {/* ── KPI row ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12}}>
        {[
          {label:"Dossiers actifs",    v:stats.total,         sub:"acquéreurs",                 color:DS.blue,   bg:DS.blueBg,   bd:DS.blueBd},
          {label:"Total encaissé",     v:F.ar(stats.rec),     sub:F.ar(stats.att)+" attendu",   color:DS.green,  bg:DS.greenBg,  bd:DS.greenBd},
          {label:"Solde à recouvrer",  v:F.ar(stats.sol),     sub:"capital restant dû",         color:DS.orange, bg:DS.orangeBg, bd:DS.orangeBd},
          {label:"Taux recouvrement",  v:F.pct(stats.taux),   sub:"sur échéances dues",         color:DS.accent, bg:DS.accentLt, bd:DS.accentMd},
          {label:"Dossiers en retard", v:stats.ret,           sub:"nécessitent une action",     color:DS.red,    bg:DS.redBg,    bd:DS.redBd},
        ].map(({label,v,sub,color,bg,bd})=>(
          <div key={label} style={{background:bg,borderRadius:10,padding:"14px 16px",border:`1px solid ${bd}`}}>
            <div style={{fontSize:10,fontWeight:600,color,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:8}}>{label}</div>
            <div style={{fontSize:22,fontWeight:800,color:DS.text,letterSpacing:"-0.5px",lineHeight:1}}>{v}</div>
            <div style={{fontSize:11,color:DS.text4,marginTop:6}}>{sub}</div>
          </div>
        ))}
      </div>

      {/* ── Ligne 2 : Graphique barres + performance sites ── */}
      <div style={{display:"grid",gridTemplateColumns:"1.6fr 1fr",gap:12}}>

        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
            <div>
              <div style={{fontSize:13,fontWeight:600,color:DS.text}}>Paiements encaissés par mois</div>
              <div style={{fontSize:11,color:DS.text4,marginTop:2}}>Exercice {annee}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:11,color:DS.text4}}>Cumul {annee}</div>
              <div style={{fontSize:16,fontWeight:700,color:DS.green}}>{F.ar(pmtMois.reduce((s,v)=>s+v,0))}</div>
            </div>
          </div>
          <BarChart data={pmtMois.map((v,i)=>({label:MOIS_COURT[i],value:v,highlight:i===moisVue}))} height={120}/>
        </Card>

        <Card>
          <div style={{fontSize:13,fontWeight:600,color:DS.text,marginBottom:14}}>Performance par site</div>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            {perfSites.map(({sigle,cfg,vendu,total,rec,att,taux,al})=>(
              <div key={sigle}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:5,alignItems:"center"}}>
                  <div style={{display:"flex",alignItems:"center",gap:7}}>
                    <div style={{width:8,height:8,borderRadius:"50%",background:cfg.couleur||DS.accent,flexShrink:0}}/>
                    <span style={{fontSize:12,fontWeight:500,color:DS.text2}}>{cfg.nom||sigle}</span>
                    {al>0&&<span style={{background:DS.redBg,color:DS.red,fontSize:9,padding:"1px 5px",borderRadius:3,fontWeight:600,border:`1px solid ${DS.redBd}`}}>{al}</span>}
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:11,color:DS.text4}}>{vendu}/{total}</span>
                    <span style={{fontSize:12,fontWeight:700,color:taux>=80?DS.green:taux>=50?DS.orange:DS.red}}>{taux.toFixed(0)}%</span>
                  </div>
                </div>
                <ProgressBar pct={taux} color={cfg.couleur||DS.accent} height={4}/>
              </div>
            ))}
          </div>
        </Card>
      </div>

      {/* ── Ligne 3 : Statuts + Taux + Semaines ── */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>

        {/* Donut statuts */}
        <Card>
          <div style={{fontSize:13,fontWeight:600,color:DS.text,marginBottom:14}}>Situation des contrats</div>
          <div style={{display:"flex",gap:14,alignItems:"center",marginBottom:14}}>
            <DonutChart size={84} segments={donutSegs}/>
            <div>
              <div style={{fontSize:22,fontWeight:800,color:DS.green,lineHeight:1}}>{stats.reg}</div>
              <div style={{fontSize:11,color:DS.text4,marginBottom:6}}>réguliers</div>
              <div style={{fontSize:18,fontWeight:700,color:DS.red,lineHeight:1}}>{stats.ret}</div>
              <div style={{fontSize:11,color:DS.text4}}>en retard</div>
            </div>
          </div>
          {Object.entries(repStatut).filter(([k])=>k!=="inactif"&&repStatut[k]>0).sort(([a],[b])=>(STATUT_ORD[a]||9)-(STATUT_ORD[b]||9)).map(([s,n])=>(
            <div key={s} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:STATUT[s]?.dot||DS.text4}}/>
                <span style={{fontSize:11,color:DS.text2}}>{STATUT[s]?.label||s}</span>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:55,background:"#f3f4f6",borderRadius:999,height:3,overflow:"hidden"}}>
                  <div style={{width:`${(n/stats.total)*100}%`,height:"100%",background:STATUT[s]?.dot||DS.text4,borderRadius:999}}/>
                </div>
                <span style={{fontSize:11,color:DS.text3,minWidth:18,textAlign:"right"}}>{n}</span>
              </div>
            </div>
          ))}
        </Card>

        {/* Taux recouvrement */}
        <Card>
          <div style={{fontSize:13,fontWeight:600,color:DS.text,marginBottom:4}}>Taux de recouvrement</div>
          <div style={{fontSize:11,color:DS.text4,marginBottom:16}}>Paiements reçus vs attendus</div>
          <div style={{textAlign:"center",marginBottom:16}}>
            <div style={{fontSize:48,fontWeight:900,color:stats.taux>=80?DS.green:stats.taux>=50?DS.orange:DS.red,lineHeight:1,letterSpacing:"-2px"}}>
              {Math.round(stats.taux)}<span style={{fontSize:22,fontWeight:600}}>%</span>
            </div>
            <div style={{fontSize:11,color:DS.text4,marginTop:4}}>{F.ar(stats.rec)} encaissé</div>
          </div>
          <ProgressBar pct={stats.taux} color={stats.taux>=80?DS.green:stats.taux>=50?DS.orange:DS.red} height={8}/>
          <div style={{display:"flex",justifyContent:"space-between",marginTop:6,fontSize:10,color:DS.text4}}>
            <span>Reçu : {F.ar(stats.rec)}</span>
            <span>Attendu : {F.ar(stats.att)}</span>
          </div>
          <div style={{marginTop:14,paddingTop:12,borderTop:`1px solid ${DS.border}`}}>
            <SectionLabel>Par site</SectionLabel>
            {perfSites.slice(0,5).map(({sigle,cfg,taux})=>(
              <div key={sigle} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <div style={{width:6,height:6,borderRadius:"50%",background:cfg.couleur||DS.accent,flexShrink:0}}/>
                <span style={{fontSize:11,color:DS.text2,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cfg.nom||sigle}</span>
                <span style={{fontSize:11,fontWeight:600,color:taux>=80?DS.green:taux>=50?DS.orange:DS.red,minWidth:32,textAlign:"right"}}>{taux.toFixed(0)}%</span>
              </div>
            ))}
          </div>
        </Card>

        {/* Évolution + semaines */}
        <Card>
          <div style={{fontSize:13,fontWeight:600,color:DS.text,marginBottom:4}}>Évolution mensuelle</div>
          <div style={{fontSize:11,color:DS.text4,marginBottom:10}}>Tendance {annee}</div>
          <div style={{marginBottom:16}}>
            <MiniLineChart data={pmtMois.map(v=>({value:v}))} height={60} color={DS.accent}/>
          </div>
          <SectionLabel>Semaines — {MOIS[moisVue]}</SectionLabel>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
            {pmtSem.map(({label,value,nb})=>(
              <div key={label} style={{background:value>0?DS.accentLt:"#f9fafb",border:`1px solid ${value>0?DS.accentMd:DS.border}`,borderRadius:7,padding:"9px 11px"}}>
                <div style={{fontSize:10,color:DS.text4,marginBottom:2}}>{label}</div>
                <div style={{fontSize:13,fontWeight:700,color:value>0?DS.accent:DS.text4}}>{value>0?F.ar(value):"—"}</div>
                <div style={{fontSize:10,color:DS.text4}}>{nb} pmt{nb!==1?"s":""}</div>
              </div>
            ))}
          </div>
        </Card>
      </div>

      {/* ── Ligne 4 : Tableau sites + Alertes ── */}
      <div style={{display:"grid",gridTemplateColumns:"1.5fr 1fr",gap:12}}>

        <Card padding="0">
          <div style={{padding:"14px 18px",borderBottom:`1px solid ${DS.border}`}}>
            <div style={{fontSize:13,fontWeight:600,color:DS.text}}>Récapitulatif financier par site</div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:"#f9fafb",borderBottom:`1px solid ${DS.border}`}}>
                {["Site","Dossiers","Encaissé","Solde restant","Taux","Alertes"].map(h=>(
                  <th key={h} style={{padding:"8px 14px",textAlign:"left",fontSize:10,fontWeight:600,color:DS.text4,textTransform:"uppercase",letterSpacing:"0.6px"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {perfSites.map(({sigle,cfg,vendu,total,rec,sol,taux,al},i)=>(
                <tr key={sigle} style={{borderBottom:`1px solid ${DS.border}`,background:i%2===0?"#fff":"#fafafa"}}>
                  <td style={{padding:"10px 14px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:8,height:8,borderRadius:"50%",background:cfg.couleur||DS.accent,flexShrink:0}}/>
                      <div>
                        <div style={{fontWeight:500,color:DS.text,fontSize:12}}>{cfg.nom||sigle}</div>
                        {cfg.region&&<div style={{fontSize:10,color:DS.text4}}>{cfg.region}</div>}
                      </div>
                    </div>
                  </td>
                  <td style={{padding:"10px 14px",color:DS.text3}}>{vendu}/{total}</td>
                  <td style={{padding:"10px 14px",color:DS.green,fontWeight:600}}>{F.ar(rec)}</td>
                  <td style={{padding:"10px 14px",color:DS.orange}}>{F.ar(sol)}</td>
                  <td style={{padding:"10px 14px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <div style={{width:44,background:"#f3f4f6",borderRadius:999,height:4,overflow:"hidden"}}>
                        <div style={{width:`${Math.min(100,taux)}%`,height:"100%",background:taux>=80?DS.green:taux>=50?DS.orange:DS.red,borderRadius:999}}/>
                      </div>
                      <span style={{fontSize:11,fontWeight:700,color:taux>=80?DS.green:taux>=50?DS.orange:DS.red}}>{taux.toFixed(0)}%</span>
                    </div>
                  </td>
                  <td style={{padding:"10px 14px"}}>
                    {al>0?<span style={{background:DS.redBg,color:DS.red,border:`1px solid ${DS.redBd}`,fontSize:10,padding:"2px 7px",borderRadius:4,fontWeight:600}}>{al}</span>:<span style={{color:DS.green,fontSize:12}}>—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>

        <Card padding="0">
          <div style={{padding:"14px 18px",borderBottom:`1px solid ${DS.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div style={{fontSize:13,fontWeight:600,color:DS.text}}>Alertes prioritaires</div>
            {alertes.length>0&&<span style={{background:DS.redBg,color:DS.red,border:`1px solid ${DS.redBd}`,fontSize:10,padding:"2px 8px",borderRadius:4,fontWeight:600}}>{alertes.length}</span>}
          </div>
          {alertes.length===0?(
            <div style={{padding:32,textAlign:"center"}}>
              <div style={{fontSize:28,marginBottom:8,color:DS.green}}>✓</div>
              <div style={{fontSize:12,color:DS.text3,fontWeight:500}}>Aucune alerte</div>
              <div style={{fontSize:11,color:DS.text4,marginTop:4}}>Tous les paiements sont à jour</div>
            </div>
          ):(
            <div style={{maxHeight:300,overflowY:"auto"}}>
              {alertes.map((a,i)=>{
                const s=getStatut(a),cfg=sitesConfig[a.sigle]||{};
                const att=calcAttendues(a),retard=Math.max(0,att-a.mensualites_ecoulees);
                const sc=STATUT[s]?.dot||DS.text4;
                return(
                  <div key={a.id} onClick={()=>onFiche(a.id)} style={{display:"flex",alignItems:"center",gap:10,padding:"11px 18px",borderBottom:i<alertes.length-1?`1px solid ${DS.border}`:"none",cursor:"pointer",background:"#fff"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#f9fafb"}
                    onMouseLeave={e=>e.currentTarget.style.background="#fff"}>
                    <div style={{width:3,height:36,background:sc,borderRadius:2,flexShrink:0}}/>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,fontWeight:500,color:DS.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.nom}</div>
                      <div style={{fontSize:10,color:DS.text4,marginTop:1}}>
                        {cfg.nom||a.sigle} · Lgt {a.num_lgt}
                        {retard>0&&<span style={{color:sc}}> · {retard} mens. dues</span>}
                      </div>
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:10,color:DS.text4}}>Dû</div>
                      <div style={{fontSize:12,fontWeight:700,color:sc}}>{retard>0?F.ar(retard*a.mensualite):"—"}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// LISTE ACQUÉREURS
// ═══════════════════════════════════════════════════════════════════
function ListeView({ actifs, filteredList, searchTerm, setSearchTerm, filterSit, setFilterSit, siteActif, sitesConfig, getStatut, STATUT, onFiche, onPmt, onEdit }) {
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div>
          <h2 style={{margin:0,fontSize:16,fontWeight:700,color:DS.text}}>{siteActif==="tous"?"Tous les acquéreurs":`Acquéreurs — ${sitesConfig[siteActif]?.nom}`}</h2>
          <div style={{fontSize:11,color:DS.text4,marginTop:3}}>{filteredList.length} dossier{filteredList.length>1?"s":""}</div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <div style={{position:"relative"}}>
            <Input value={searchTerm} onChange={e=>setSearchTerm(e.target.value)} placeholder="Rechercher..." style={{paddingLeft:30,width:220}}/>
            <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:DS.text4,fontSize:13,pointerEvents:"none"}}>⌕</span>
          </div>
          <select value={filterSit} onChange={e=>setFilterSit(e.target.value)} style={{background:"#fff",border:`1px solid ${DS.border2}`,borderRadius:6,padding:"8px 10px",fontSize:12,color:DS.text2,cursor:"pointer",outline:"none"}}>
            <option value="tous">Toutes situations</option>
            <option value="légalisé">Légalisé</option>
            <option value="signé">Signé</option>
            <option value="suspendu">Suspendu</option>
          </select>
        </div>
      </div>
      <Card padding="0">
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:900}}>
          <thead>
            <tr style={{background:"#f9fafb",borderBottom:`1px solid ${DS.border}`}}>
              {["Site","N° Lgt","Nom et Prénoms","Contact","Situation","Payé","Solde","Mens.","Statut",""].map(h=>(
                <th key={h} style={{padding:"9px 13px",textAlign:"left",fontSize:10,fontWeight:600,color:DS.text4,textTransform:"uppercase",letterSpacing:"0.6px"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredList.map((a,i)=>{
              const s=getStatut(a),cfg=sitesConfig[a.sigle]||{};
              return(
                <tr key={a.id} style={{borderBottom:`1px solid ${DS.border}`,background:"#fff"}}
                  onMouseEnter={e=>e.currentTarget.style.background="#f9fafb"}
                  onMouseLeave={e=>e.currentTarget.style.background="#fff"}>
                  <td style={{padding:"9px 13px"}}>
                    <span style={{background:(cfg.couleur||DS.accent)+"18",color:cfg.couleur||DS.accent,fontSize:10,padding:"2px 7px",borderRadius:4,fontWeight:700,border:`1px solid ${(cfg.couleur||DS.accent)}33`}}>{a.sigle}</span>
                  </td>
                  <td style={{padding:"9px 13px",fontSize:12,fontWeight:600,color:DS.text}}>{a.num_lgt}</td>
                  <td style={{padding:"9px 13px",fontSize:12,color:DS.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.nom}</td>
                  <td style={{padding:"9px 13px",fontSize:11,color:DS.text3}}>{a.contact}</td>
                  <td style={{padding:"9px 13px"}}><span style={{background:"#f3f4f6",color:DS.text3,fontSize:10,padding:"2px 7px",borderRadius:4,border:`1px solid ${DS.border}`}}>{a.situation}</span></td>
                  <td style={{padding:"9px 13px",fontSize:11,fontWeight:600,color:DS.green,whiteSpace:"nowrap"}}>{F.ar(a.loyer_paye)}</td>
                  <td style={{padding:"9px 13px",fontSize:11,color:DS.orange,whiteSpace:"nowrap"}}>{F.ar(a.solde)}</td>
                  <td style={{padding:"9px 13px",fontSize:11,color:DS.text4}}>{a.mensualites_ecoulees}/{a.nb_mensualites}</td>
                  <td style={{padding:"9px 13px"}}><StatutBadge statut={s} STATUT={STATUT}/></td>
                  <td style={{padding:"9px 13px"}}>
                    <div style={{display:"flex",gap:5}}>
                      <Btn onClick={()=>onFiche(a.id)} size="sm" variant="ghost">Fiche</Btn>
                      <Btn onClick={()=>onPmt(a.id)} size="sm" variant="outline" color={DS.green}>Paiement</Btn>
                      <Btn onClick={()=>onEdit(a)} size="sm" variant="ghost">Modifier</Btn>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {filteredList.length===0&&<div style={{padding:40,textAlign:"center",color:DS.text4,fontSize:13}}>Aucun résultat</div>}
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// ALERTES
// ═══════════════════════════════════════════════════════════════════
function AlertesView({ alertes, getStatut, STATUT, sitesConfig, calcAttendues, onPmt, onFiche, onEdit }) {
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div>
          <h2 style={{margin:0,fontSize:16,fontWeight:700,color:DS.text}}>Alertes & Relances</h2>
          <div style={{fontSize:11,color:DS.text4,marginTop:3}}>Dossiers nécessitant une action immédiate</div>
        </div>
        {alertes.length>0&&<span style={{background:DS.redBg,border:`1px solid ${DS.redBd}`,borderRadius:6,padding:"5px 12px",fontSize:12,color:DS.red,fontWeight:600}}>{alertes.length} dossier{alertes.length>1?"s":""} en attente</span>}
      </div>
      {alertes.length===0?(
        <Card style={{textAlign:"center",padding:"60px 20px"}}>
          <div style={{fontSize:36,color:DS.green,marginBottom:12}}>✓</div>
          <div style={{fontSize:14,fontWeight:600,color:DS.text}}>Aucune alerte</div>
          <div style={{fontSize:12,color:DS.text4,marginTop:6}}>Tous les paiements sont à jour</div>
        </Card>
      ):(
        <Card padding="0">
          {alertes.map((a,i)=>{
            const s=getStatut(a),cfg=sitesConfig[a.sigle]||{};
            const att=calcAttendues(a),retard=Math.max(0,att-a.mensualites_ecoulees);
            const sc=s==="urgente"?DS.red:s==="retard"?DS.orange:"#ca8a04";
            return(
              <div key={a.id} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",borderBottom:i<alertes.length-1?`1px solid ${DS.border}`:"none",background:"#fff"}}
                onMouseEnter={e=>e.currentTarget.style.background="#f9fafb"}
                onMouseLeave={e=>e.currentTarget.style.background="#fff"}>
                <div style={{width:3,height:44,background:sc,borderRadius:2,flexShrink:0}}/>
                <div style={{display:"flex",alignItems:"center",gap:7,minWidth:120}}>
                  <span style={{background:(cfg.couleur||DS.accent)+"18",color:cfg.couleur||DS.accent,fontSize:9,padding:"2px 6px",borderRadius:3,fontWeight:700}}>{a.sigle}</span>
                  <StatutBadge statut={s} STATUT={STATUT}/>
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:12,fontWeight:500,color:DS.text}}>{a.nom}</div>
                  <div style={{fontSize:10,color:DS.text4,marginTop:2}}>
                    Logement {a.num_lgt} · {a.contact}
                    {retard>0&&<span style={{color:sc}}> · {retard} mensualité{retard>1?"s":""} dues</span>}
                    {s==="urgente"&&<span style={{color:sc}}> · Aucun paiement depuis la signature</span>}
                  </div>
                </div>
                <div style={{textAlign:"right",minWidth:100}}>
                  <div style={{fontSize:10,color:DS.text4}}>Montant dû</div>
                  <div style={{fontSize:13,fontWeight:700,color:sc}}>{retard>0?F.full(retard*a.mensualite):"—"}</div>
                </div>
                <div style={{display:"flex",gap:6}}>
                  <Btn onClick={()=>onPmt(a.id)} variant="solid" color={DS.accent} size="sm">+ Paiement</Btn>
                  <Btn onClick={()=>onFiche(a.id)} variant="ghost" size="sm">Fiche</Btn>
                </div>
              </div>
            );
          })}
        </Card>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// FICHE DÉTAILLÉE
// ═══════════════════════════════════════════════════════════════════
function FicheDetaillee({ a, onPaiement, onPrint, onDelete, onEdit, getStatut, STATUT, calcAttendues, onBack, sitesConfig }) {
  const s=getStatut(a),cfg=sitesConfig[a.sigle]||{},att=calcAttendues(a);
  const pct=a.prix_logement>0?(a.loyer_paye/a.prix_logement)*100:0;
  const sc=STATUT[s]?.dot||DS.text4;
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
        <Btn onClick={onBack} variant="ghost" size="sm">← Retour</Btn>
        <span style={{background:(cfg.couleur||DS.accent)+"18",color:cfg.couleur||DS.accent,fontSize:11,padding:"2px 8px",borderRadius:4,fontWeight:600,border:`1px solid ${(cfg.couleur||DS.accent)}33`}}>{cfg.nom||a.sigle}</span>
        <h2 style={{margin:0,fontSize:16,fontWeight:700,color:DS.text,flex:1}}>{a.nom}</h2>
        <StatutBadge statut={s} STATUT={STATUT}/>
        <Btn onClick={onEdit}     variant="ghost"   size="sm">Modifier</Btn>
        <Btn onClick={onPaiement} variant="solid"   size="sm" color={DS.accent}>+ Paiement</Btn>
        <Btn onClick={onPrint}    variant="outline"  size="sm" color={DS.blue}>Imprimer</Btn>
      </div>

      {/* Progression */}
      <Card style={{marginBottom:12}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:14}}>
          {[["Prix logement",F.full(a.prix_logement),DS.text2],["Total payé",F.full(a.loyer_paye),DS.green],["Solde restant",F.full(a.solde),DS.orange],[`Mensualités (${att} attendues)`,`${a.mensualites_ecoulees}/${a.nb_mensualites}`,a.mensualites_ecoulees>=att?DS.green:DS.orange]].map(([l,v,c])=>(
            <div key={l} style={{background:"#f9fafb",borderRadius:7,padding:"10px 12px",border:`1px solid ${DS.border}`}}>
              <div style={{fontSize:10,color:DS.text4,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.4px"}}>{l}</div>
              <div style={{fontSize:15,fontWeight:700,color:c}}>{v}</div>
            </div>
          ))}
        </div>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:DS.text4,marginBottom:5}}>
          <span>Progression remboursement</span>
          <span style={{fontWeight:600,color:sc}}>{pct.toFixed(1)}%</span>
        </div>
        <ProgressBar pct={pct} color={sc} height={8}/>
      </Card>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
        {[
          {title:"Informations personnelles",rows:[["Nom",a.nom],["CIN",a.cin||"—"],["Contact",a.contact||"—"],["Correspondance",a.correspondance||"—"],["Adresse",a.adresse||"—"],["Situation",a.situation]]},
          {title:"Détails du contrat",rows:[["Réf. contrat",a.id],["Date signature",a.date_signature||"—"],["Date fin",a.date_fin||"—"],["Durée",`${a.nb_mensualites} mois`],["Mensualité",F.full(a.mensualite)],["Acompte",F.full(a.acompte||0)]]},
        ].map(({title,rows})=>(
          <Card key={title}>
            <SectionLabel>{title}</SectionLabel>
            {rows.map(([l,v])=>(
              <div key={l} style={{display:"flex",padding:"5px 0",borderBottom:`1px solid ${DS.border}`}}>
                <span style={{width:140,fontSize:12,color:DS.text4,flexShrink:0}}>{l}</span>
                <span style={{fontSize:12,color:DS.text,fontWeight:400}}>{v}</span>
              </div>
            ))}
          </Card>
        ))}
      </div>

      <Card>
        <SectionLabel>Historique des paiements ({(a.paiements||[]).length})</SectionLabel>
        {(a.paiements||[]).length===0?(
          <div style={{padding:"16px 0",textAlign:"center",color:DS.text4,fontSize:12}}>Aucun paiement enregistré</div>
        ):(
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr style={{borderBottom:`1px solid ${DS.border}`}}>{["Date","Montant","Réf. Quittance","Mode","Note",""].map(h=><th key={h} style={{padding:"6px 10px",fontSize:10,fontWeight:600,color:DS.text4,textAlign:"left",textTransform:"uppercase",letterSpacing:"0.5px"}}>{h}</th>)}</tr></thead>
            <tbody>{(a.paiements||[]).map((p,i)=>(
              <tr key={p.id} style={{borderBottom:`1px solid ${DS.border}`,background:i%2===0?"#fff":"#fafafa"}}>
                <td style={{padding:"7px 10px",fontSize:12,color:DS.text3}}>{p.date}</td>
                <td style={{padding:"7px 10px",fontSize:12,fontWeight:600,color:DS.green}}>{F.full(p.montant)}</td>
                <td style={{padding:"7px 10px",fontSize:12,color:DS.blue}}>{p.ref_quittance||"—"}</td>
                <td style={{padding:"7px 10px",fontSize:12,color:DS.text3}}>{p.mode}</td>
                <td style={{padding:"7px 10px",fontSize:12,color:DS.text4}}>{p.note||"—"}</td>
                <td style={{padding:"7px 10px"}}><Btn onClick={()=>onDelete(a.id,p.id)} variant="danger" size="sm">Supprimer</Btn></td>
              </tr>
            ))}</tbody>
          </table>
        )}
        {a.note&&<div style={{marginTop:10,padding:"8px 12px",background:DS.orangeBg,border:`1px solid ${DS.orangeBd}`,borderRadius:6,fontSize:12,color:DS.orange}}>Note : {a.note}</div>}
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// FICHE IMPRIMABLE
// ═══════════════════════════════════════════════════════════════════
function FicheImprimable({ a, onClose, fmt, calcAttendues, getStatut, STATUT, sitesConfig }) {
  const s=getStatut(a),cfg=sitesConfig[a.sigle]||{},att=calcAttendues(a),pct=a.prix_logement>0?(a.loyer_paye/a.prix_logement)*100:0;
  useEffect(()=>{setTimeout(()=>window.print(),400);},[]);
  return (
    <div style={{fontFamily:"Georgia,serif",background:"#fff",color:"#111",padding:32,maxWidth:800,margin:"0 auto"}}>
      <style>{`@media print{button{display:none!important}}`}</style>
      <button onClick={onClose} style={{background:"#f3f4f6",border:"1px solid #e5e7eb",borderRadius:4,padding:"4px 10px",cursor:"pointer",marginBottom:16,fontSize:11}}>Fermer</button>
      <div style={{borderBottom:"2px solid #B5006E",paddingBottom:14,marginBottom:18,display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
        <div><img src={LOGO_SRC} alt="ANALOGH" style={{height:36,objectFit:"contain"}}/><div style={{fontSize:10,color:"#6b7280",marginTop:4}}>{cfg.nom||a.sigle} · Suivi Remboursement</div></div>
        <div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:700,color:"#B5006E"}}>FICHE CLIENT</div><div style={{fontSize:10,color:"#6b7280"}}>{new Date().toLocaleDateString("fr-FR")}</div></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
        {[{t:"Identité",rows:[["Nom",a.nom],["Lgt",a.num_lgt],["Contact",a.contact],["CIN",a.cin||"—"],["Situation",a.situation]]},{t:"Contrat",rows:[["Réf.",a.id],["Prix",fmt(a.prix_logement)],["Mensualité",fmt(a.mensualite)],["Durée",`${a.nb_mensualites} mois`],["Signature",a.date_signature||"—"]]}].map(({t,rows})=>(
          <div key={t} style={{border:"1px solid #e5e7eb",borderRadius:4,padding:12}}>
            <div style={{fontSize:9,fontWeight:700,color:"#B5006E",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>{t}</div>
            {rows.map(([l,v])=>(<div key={l} style={{display:"flex",paddingBottom:3,marginBottom:3,borderBottom:"1px solid #f5f5f5"}}><span style={{width:110,fontSize:9,color:"#9ca3af",flexShrink:0}}>{l} :</span><span style={{fontSize:10}}>{v}</span></div>))}
          </div>
        ))}
      </div>
      <div style={{border:"2px solid #B5006E",borderRadius:4,padding:12,marginBottom:14}}>
        <div style={{fontSize:9,fontWeight:700,color:"#B5006E",marginBottom:8,textTransform:"uppercase"}}>Remboursement</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:10}}>
          {[["Payé",fmt(a.loyer_paye),"#16a34a"],["Solde",fmt(a.solde),"#d97706"],["Mens.",`${a.mensualites_ecoulees}/${a.nb_mensualites}`,"#2563eb"],["Attendues",att,a.mensualites_ecoulees>=att?"#16a34a":"#dc2626"]].map(([l,v,c])=>(
            <div key={l} style={{background:"#f9fafb",borderRadius:4,padding:7,border:"1px solid #e5e7eb"}}><div style={{fontSize:7,color:"#9ca3af"}}>{l}</div><div style={{fontSize:12,fontWeight:700,color:c}}>{v}</div></div>
          ))}
        </div>
        <div style={{background:"#e5e7eb",borderRadius:999,height:7}}><div style={{width:`${Math.min(100,pct)}%`,height:"100%",background:"#B5006E",borderRadius:999}}/></div>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:2,fontSize:8,color:"#9ca3af"}}><span>0%</span><span style={{fontWeight:700,color:"#111"}}>{pct.toFixed(1)}%</span><span>100%</span></div>
      </div>
      {(a.paiements||[]).length>0&&(<div style={{border:"1px solid #e5e7eb",borderRadius:4,padding:12,marginBottom:14}}>
        <div style={{fontSize:9,fontWeight:700,color:"#B5006E",marginBottom:8,textTransform:"uppercase"}}>Historique</div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:9}}>
          <thead><tr style={{background:"#fdf2f8"}}>{["Date","Montant","Réf.","Mode"].map(h=><th key={h} style={{padding:"3px 6px",textAlign:"left",borderBottom:"1px solid #e5e7eb",color:"#B5006E"}}>{h}</th>)}</tr></thead>
          <tbody>{(a.paiements||[]).map(p=>(<tr key={p.id} style={{borderBottom:"1px solid #f5f5f5"}}><td style={{padding:"3px 6px"}}>{p.date}</td><td style={{padding:"3px 6px",fontWeight:700,color:"#16a34a"}}>{fmt(p.montant)}</td><td style={{padding:"3px 6px"}}>{p.ref_quittance||"—"}</td><td style={{padding:"3px 6px"}}>{p.mode}</td></tr>))}</tbody>
        </table>
      </div>)}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:48,marginTop:24,paddingTop:14,borderTop:"1px solid #e5e7eb"}}>
        {["Signature du Gestionnaire","Signature de l'Acquéreur"].map(sig=>(<div key={sig} style={{textAlign:"center"}}><div style={{height:40}}/><div style={{borderTop:"1px solid #111",paddingTop:4,fontSize:8,color:"#9ca3af"}}>{sig}</div></div>))}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// SAISIE PAIEMENT
// ═══════════════════════════════════════════════════════════════════
function SaisieView({ acquéreurs, onSave, onFiche, onEdit, getStatut, STATUT, sitesConfig }) {
  const [selId,setSelId]=useState(""),  [search,setSearch]=useState("");
  const [form,setForm]=useState({montant:"",date:new Date().toISOString().split("T")[0],ref_quittance:"",mode:"virement",note:""});
  const [ok,setOk]=useState(false);
  const sel=acquéreurs.find(a=>a.id===selId);
  const filtered=acquéreurs.filter(a=>isActif(a)&&(!search||a.nom.toLowerCase().includes(search.toLowerCase())||a.num_lgt.includes(search)||(a.sigle||"").toLowerCase().includes(search.toLowerCase())));
  function save(){if(!sel||!form.montant) return;onSave(sel.id,{...form,montant:parseFloat(form.montant)});setOk(true);setForm(f=>({...f,montant:"",ref_quittance:"",note:""}));setTimeout(()=>setOk(false),3000);}
  return(
    <div style={{display:"grid",gridTemplateColumns:"280px 1fr",gap:16}}>
      <Card padding="0">
        <div style={{padding:"12px 14px",borderBottom:`1px solid ${DS.border}`}}>
          <div style={{fontSize:12,fontWeight:600,color:DS.text,marginBottom:8}}>Sélectionner</div>
          <div style={{position:"relative"}}><Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Rechercher..." style={{paddingLeft:28}}/><span style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",color:DS.text4,fontSize:13,pointerEvents:"none"}}>⌕</span></div>
        </div>
        <div style={{maxHeight:480,overflowY:"auto"}}>
          {filtered.map(a=>{const s=getStatut(a),cfg=sitesConfig[a.sigle]||{};return(
            <div key={a.id} onClick={()=>{setSelId(a.id);setForm(f=>({...f,montant:a.mensualite}));}} style={{padding:"10px 14px",cursor:"pointer",borderLeft:`3px solid ${selId===a.id?DS.accent:"transparent"}`,background:selId===a.id?"#fdf2f8":"#fff",borderBottom:`1px solid ${DS.border}`}}
              onMouseEnter={e=>{if(selId!==a.id)e.currentTarget.style.background="#f9fafb";}}
              onMouseLeave={e=>{if(selId!==a.id)e.currentTarget.style.background="#fff";}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:11,fontWeight:500,color:DS.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:170}}>{a.nom}</span>
                <div style={{width:6,height:6,borderRadius:"50%",background:STATUT[s]?.dot||DS.text4,flexShrink:0}}/>
              </div>
              <div style={{fontSize:10,color:DS.text4,marginTop:2}}><span style={{color:cfg.couleur||DS.accent,fontWeight:600}}>{a.sigle}</span>{" · "}Lgt {a.num_lgt}</div>
            </div>
          );})}
          {filtered.length===0&&<div style={{padding:24,textAlign:"center",color:DS.text4,fontSize:12}}>Aucun résultat</div>}
        </div>
      </Card>
      <div>
        {!sel?(
          <Card style={{textAlign:"center",padding:"60px 20px",border:`1px dashed ${DS.border2}`}}>
            <div style={{fontSize:12,color:DS.text4}}>Sélectionnez un acquéreur dans la liste</div>
          </Card>
        ):(
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <Card>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div><div style={{fontSize:14,fontWeight:700,color:DS.text}}>{sel.nom}</div><div style={{fontSize:11,color:DS.text4,marginTop:2}}>Lgt {sel.num_lgt} · {sel.contact} · {(sitesConfig[sel.sigle]||{}).nom}</div></div>
                <div style={{display:"flex",gap:6}}><Btn onClick={()=>onEdit(sel)} variant="ghost" size="sm">Modifier</Btn><Btn onClick={()=>onFiche(sel.id)} variant="ghost" size="sm">Fiche</Btn></div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                {[["Déjà payé",F.full(sel.loyer_paye),DS.green],["Solde restant",F.full(sel.solde),DS.orange],["Mensualités",`${sel.mensualites_ecoulees}/${sel.nb_mensualites}`,DS.text2]].map(([l,v,c])=>(
                  <div key={l} style={{background:"#f9fafb",borderRadius:6,padding:"8px 10px",border:`1px solid ${DS.border}`}}><div style={{fontSize:10,color:DS.text4}}>{l}</div><div style={{fontSize:13,fontWeight:700,color:c}}>{v}</div></div>
                ))}
              </div>
            </Card>
            <Card>
              <div style={{fontSize:13,fontWeight:600,color:DS.text,marginBottom:14}}>Enregistrer un paiement</div>
              {ok&&<div style={{background:DS.greenBg,border:`1px solid ${DS.greenBd}`,borderRadius:6,padding:"8px 12px",marginBottom:12,color:DS.green,fontSize:12}}>Paiement enregistré avec succès</div>}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                {[{l:"Montant (Ar) *",k:"montant",t:"number"},{l:"Date *",k:"date",t:"date"},{l:"Réf. Quittance",k:"ref_quittance",t:"text",p:"DR N°..."}].map(({l,k,t,p})=>(
                  <div key={k}><label style={MODAL_LBL}>{l}</label><Input type={t} value={form[k]} onChange={e=>setForm(f=>({...f,[k]:e.target.value}))} placeholder={p}/></div>
                ))}
                <div><label style={MODAL_LBL}>Mode de paiement</label>
                  <select value={form.mode} onChange={e=>setForm(f=>({...f,mode:e.target.value}))} style={{width:"100%",background:"#fff",border:`1px solid ${DS.border2}`,borderRadius:6,padding:"8px 11px",fontSize:13,color:DS.text,outline:"none"}}>
                    <option value="virement">Virement permanent</option><option value="virement_p">Virement ponctuel</option><option value="espece">Espèces</option><option value="cheque">Chèque de banque</option><option value="retenu">Retenu à la source</option>
                  </select>
                </div>
              </div>
              <div style={{marginTop:10}}><label style={MODAL_LBL}>Note</label><Input value={form.note} onChange={e=>setForm(f=>({...f,note:e.target.value}))} placeholder="Observation..."/></div>
              <div style={{marginTop:14}}>
                <Btn onClick={save} variant="solid" color={DS.accent} disabled={!form.montant} xtra={{width:"100%",display:"block",textAlign:"center",boxShadow:form.montant?`0 4px 14px ${DS.accent}33`:"none"}}>
                  Enregistrer{form.montant?` — ${F.full(parseFloat(form.montant)||0)}`:""}
                </Btn>
              </div>
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MODAL PAIEMENT
// ═══════════════════════════════════════════════════════════════════
// ─── Styles partagés modals (stables entre renders) ───
const MODAL_SI  = {background:"#fff",border:`1px solid ${DS.border2}`,borderRadius:6,padding:"8px 11px",fontSize:13,color:DS.text,width:"100%",boxSizing:"border-box",outline:"none"};
const MODAL_LBL = {display:"block",fontSize:11,fontWeight:500,color:DS.text2,marginBottom:5};

function ModalPaiement({ a, onSave, onClose }) {
  const [form,setForm]=useState({montant:a.mensualite||670000,date:new Date().toISOString().split("T")[0],ref_quittance:"",mode:"virement",note:""});
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(17,24,39,0.4)",backdropFilter:"blur(4px)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}} onMouseDown={e=>{if(e.target===e.currentTarget)onClose();}}>
      <Card style={{width:440,boxShadow:DS.shadowLg}} onMouseDown={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontSize:15,fontWeight:700,color:DS.text}}>Saisir un paiement</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:DS.text4,cursor:"pointer",fontSize:20,lineHeight:1}}>×</button>
        </div>
        <div style={{background:"#f9fafb",borderRadius:7,padding:"10px 12px",marginBottom:14,border:`1px solid ${DS.border}`,fontSize:12}}>
          <span style={{fontWeight:600,color:DS.text}}>{a.nom}</span> · Lgt {a.num_lgt}
          <span style={{float:"right",color:DS.orange}}>Solde : {F.full(a.solde)}</span>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {[{l:"Montant (Ar) *",k:"montant",t:"number"},{l:"Date *",k:"date",t:"date"},{l:"Réf. Quittance",k:"ref_quittance",t:"text",p:"DR N°..."}].map(({l,k,t,p})=>(
            <div key={k}><label style={MODAL_LBL}>{l}</label><Input type={t} value={form[k]} onChange={e=>setForm(f=>({...f,[k]:e.target.value}))} placeholder={p}/></div>
          ))}
          <div><label style={MODAL_LBL}>Mode</label>
            <select value={form.mode} onChange={e=>setForm(f=>({...f,mode:e.target.value}))} style={{width:"100%",background:"#fff",border:`1px solid ${DS.border2}`,borderRadius:6,padding:"8px 11px",fontSize:13,color:DS.text,outline:"none",boxSizing:"border-box"}}>
              <option value="virement">Virement permanent</option><option value="virement_p">Virement ponctuel</option><option value="espece">Espèces</option><option value="cheque">Chèque de banque</option><option value="retenu">Retenu à la source</option>
            </select>
          </div>
          <div><label style={MODAL_LBL}>Note</label><Input value={form.note} onChange={e=>setForm(f=>({...f,note:e.target.value}))} placeholder="Observation..."/></div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16}}>
          <Btn onClick={onClose} variant="ghost" style={{flex:1}}>Annuler</Btn>
          <Btn onClick={()=>onSave(a.id,{...form,montant:parseFloat(form.montant)})} variant="solid" color={DS.accent} style={{flex:2}}>
            Enregistrer{form.montant?` — ${F.full(parseFloat(form.montant)||0)}`:""}
          </Btn>
        </div>
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MODAL ACQUÉREUR
// ═══════════════════════════════════════════════════════════════════
function ModalAcquereur({ sitesConfig, onSave, onClose, mode, initial }) {
  const isEdit=mode==="edit",sitesDispos=Object.keys(sitesConfig),def=isEdit?initial:{};
  const [form,setForm]=useState({nom:def.nom||"",cin:def.cin||"",adresse:def.adresse||"",contact:def.contact||"",correspondance:def.correspondance||"",num_lgt:def.num_lgt||"",sigle:def.sigle||sitesDispos[0]||"",id:def.id||"",date_signature:def.date_signature||"",date_fin:def.date_fin||"",prix_logement:def.prix_logement||"",acompte:def.acompte||"",nb_mensualites:def.nb_mensualites||60,mensualite:def.mensualite||"",situation:def.situation||"légalisé",mode_quittance:def.mode_quittance||"whatsapp",quittance_contact:def.quittance_contact||"",note:def.note||"",type_logement:def.type_logement||"",region:def.region||""});
  const cfg=sitesConfig[form.sigle]||{};
  function hc(k,v){setForm(f=>{const n={...f,[k]:v};if(k==="date_signature"||k==="nb_mensualites"){const sig=k==="date_signature"?v:f.date_signature,nb=k==="nb_mensualites"?v:f.nb_mensualites;if(sig){const d=new Date(sig);d.setMonth(d.getMonth()+parseInt(nb));n.date_fin=d.toISOString().split("T")[0];}}if(k==="sigle"){const c=sitesConfig[v]||{};if(!f.prix_logement)n.prix_logement=c.prix_defaut||"";if(!f.mensualite)n.mensualite=c.mensualite_defaut||"";if(!f.region)n.region=c.region||"";}return n;});}
  function save(){if(!form.nom||(!isEdit&&!form.num_lgt)) return;if(isEdit){onSave({nom:form.nom,cin:form.cin,adresse:form.adresse,contact:form.contact,correspondance:form.correspondance,sigle:form.sigle,situation:form.situation,date_signature:form.date_signature,date_fin:form.date_fin,prix_logement:parseFloat(form.prix_logement)||initial.prix_logement,mensualite:parseFloat(form.mensualite)||initial.mensualite,nb_mensualites:parseInt(form.nb_mensualites)||initial.nb_mensualites,mode_quittance:form.mode_quittance,quittance_contact:form.quittance_contact,note:form.note,type_logement:form.type_logement,region:form.region});}else{onSave({...form,prix_logement:parseFloat(form.prix_logement)||0,acompte:parseFloat(form.acompte)||0,mensualite:parseFloat(form.mensualite)||0,nb_mensualites:parseInt(form.nb_mensualites)||60});}}
  const SecH=({c,t})=><div style={{fontSize:10,fontWeight:700,color:c,textTransform:"uppercase",letterSpacing:"0.7px",marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${DS.border}`}}>{t}</div>;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(17,24,39,0.4)",backdropFilter:"blur(4px)",display:"flex",alignItems:"flex-start",justifyContent:"center",zIndex:1000,overflowY:"auto",padding:"20px 0"}} onMouseDown={e=>{if(e.target===e.currentTarget)onClose();}}>
      <Card style={{width:620,boxShadow:DS.shadowLg,margin:"auto"}} onMouseDown={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div><div style={{fontSize:15,fontWeight:700,color:DS.text}}>{isEdit?"Modifier le dossier":"Nouveau dossier acquéreur"}</div>{isEdit&&<div style={{fontSize:11,color:DS.text4,marginTop:2}}>{initial.nom}</div>}</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:DS.text4,cursor:"pointer",fontSize:22,lineHeight:1}}>×</button>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {/* Identité */}
          <div style={{background:"#f9fafb",borderRadius:8,padding:14,border:`1px solid ${DS.border}`}}>
            <div style={{fontSize:10,fontWeight:700,color:DS.green,textTransform:"uppercase",letterSpacing:"0.7px",marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${DS.border}`}}>Identité</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <div style={{gridColumn:"1/-1"}}><label style={MODAL_LBL}>Nom et Prénoms *</label><input value={form.nom} onChange={e=>hc("nom",e.target.value)} style={MODAL_SI}/></div>
              <div><label style={MODAL_LBL}>CIN</label><input value={form.cin} onChange={e=>hc("cin",e.target.value)} style={MODAL_SI}/></div>
              <div><label style={MODAL_LBL}>Contact</label><input value={form.contact} onChange={e=>hc("contact",e.target.value)} style={MODAL_SI}/></div>
              <div><label style={MODAL_LBL}>Correspondance</label><input value={form.correspondance} onChange={e=>hc("correspondance",e.target.value)} style={MODAL_SI}/></div>
              <div style={{gridColumn:"1/-1"}}><label style={MODAL_LBL}>Adresse</label><input value={form.adresse} onChange={e=>hc("adresse",e.target.value)} style={MODAL_SI}/></div>
            </div>
          </div>
          {/* Contrat */}
          <div style={{background:"#f9fafb",borderRadius:8,padding:14,border:`1px solid ${DS.border}`}}>
            <div style={{fontSize:10,fontWeight:700,color:DS.blue,textTransform:"uppercase",letterSpacing:"0.7px",marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${DS.border}`}}>Contrat</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
              <div><label style={MODAL_LBL}>Site *{isEdit&&<span style={{color:DS.orange,fontSize:9,marginLeft:4}}>(modifiable)</span>}</label><select value={form.sigle} onChange={e=>hc("sigle",e.target.value)} style={MODAL_SI}>{sitesDispos.map(s=><option key={s} value={s}>{sitesConfig[s]?.nom||s}</option>)}</select>{isEdit&&form.sigle!==initial?.sigle&&<div style={{fontSize:9,color:DS.orange,marginTop:2}}>{initial.sigle} → {form.sigle}</div>}</div>
              <div><label style={MODAL_LBL}>N° Logement{!isEdit&&" *"}</label><input value={form.num_lgt} onChange={e=>hc("num_lgt",e.target.value)} style={MODAL_SI}/></div>
              <div><label style={MODAL_LBL}>Situation</label><select value={form.situation} onChange={e=>hc("situation",e.target.value)} style={MODAL_SI}><option value="légalisé">Légalisé</option><option value="signé">Signé</option><option value="suspendu">Suspendu</option></select></div>
              {!isEdit&&<div><label style={MODAL_LBL}>Réf. Contrat</label><input value={form.id} onChange={e=>hc("id",e.target.value)} placeholder={`P${form.num_lgt||"?"}/${form.sigle}`} style={MODAL_SI}/></div>}
              <div><label style={MODAL_LBL}>Date signature</label><input type="date" value={form.date_signature} onChange={e=>hc("date_signature",e.target.value)} style={MODAL_SI}/></div>
              <div><label style={MODAL_LBL}>Durée (mois)</label><select value={form.nb_mensualites} onChange={e=>hc("nb_mensualites",e.target.value)} style={MODAL_SI}>{[24,36,48,60,72,84].map(n=><option key={n} value={n}>{n} mois</option>)}</select></div>
              <div><label style={MODAL_LBL}>Date fin</label><input value={form.date_fin} readOnly style={{...MODAL_SI,background:"#f3f4f6",cursor:"not-allowed",color:DS.text4}}/></div>
              <div><label style={MODAL_LBL}>Type</label><select value={form.type_logement} onChange={e=>hc("type_logement",e.target.value)} style={MODAL_SI}>{["","F2","F3","F4","F5","Studio","Villa","Appartement"].map(t=><option key={t} value={t}>{t||"—"}</option>)}</select></div>
              <div><label style={MODAL_LBL}>Région</label><input value={form.region} onChange={e=>hc("region",e.target.value)} placeholder={cfg.region||""} style={MODAL_SI}/></div>
            </div>
          </div>
          {/* Financier */}
          <div style={{background:"#f9fafb",borderRadius:8,padding:14,border:`1px solid ${DS.border}`}}>
            <div style={{fontSize:10,fontWeight:700,color:DS.orange,textTransform:"uppercase",letterSpacing:"0.7px",marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${DS.border}`}}>Financier</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
              <div><label style={MODAL_LBL}>Prix logement (Ar)</label><input type="number" value={form.prix_logement} onChange={e=>hc("prix_logement",e.target.value)} style={MODAL_SI}/></div>
              {!isEdit&&<div><label style={MODAL_LBL}>Acompte (Ar)</label><input type="number" value={form.acompte} onChange={e=>hc("acompte",e.target.value)} placeholder="0" style={MODAL_SI}/></div>}
              <div><label style={MODAL_LBL}>Mensualité (Ar)</label><input type="number" value={form.mensualite} onChange={e=>hc("mensualite",e.target.value)} style={MODAL_SI}/></div>
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:18}}>
          <Btn onClick={onClose} variant="ghost" style={{flex:1}}>Annuler</Btn>
          <Btn onClick={save} variant="solid" color={DS.accent} disabled={!form.nom||(isEdit?false:!form.num_lgt)} style={{flex:2}}>{isEdit?"Enregistrer les modifications":"Créer le dossier"}</Btn>
        </div>
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MODAL NOUVEAU SITE
// ═══════════════════════════════════════════════════════════════════
function ModalNouveauSite({ onSave, onClose }) {
  const COLS=["#B5006E","#2563eb","#16a34a","#d97706","#7c3aed","#dc2626","#0891b2","#65a30d"];
  const [form,setForm]=useState({nom:"",sigle:"",localite:"",region:"",couleur:COLS[0],prix_defaut:40000000,mensualite_defaut:670000,nb_logements_total:0});
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(17,24,39,0.4)",backdropFilter:"blur(4px)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}} onMouseDown={e=>{if(e.target===e.currentTarget)onClose();}}>
      <Card style={{width:460,boxShadow:DS.shadowLg}} onMouseDown={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
          <div style={{fontSize:15,fontWeight:700,color:DS.text}}>Nouveau site</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:DS.text4,cursor:"pointer",fontSize:22,lineHeight:1}}>×</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          {[{l:"Nom du site *",k:"nom",p:"Ex: Mahajanga"},{l:"Sigle *",k:"sigle",p:"MJG"},{l:"Localité",k:"localite",p:"Mahajanga ville"},{l:"Région",k:"region",p:"Boeny"}].map(({l,k,p})=>(
            <div key={k}><label style={MODAL_LBL}>{l}</label><input value={form[k]} onChange={e=>setForm(f=>({...f,[k]:k==="sigle"?e.target.value.toUpperCase().slice(0,4):e.target.value}))} placeholder={p} style={MODAL_SI}/></div>
          ))}
          <div><label style={MODAL_LBL}>Nb logements</label><input type="number" value={form.nb_logements_total} onChange={e=>setForm(f=>({...f,nb_logements_total:+e.target.value}))} style={MODAL_SI}/></div>
          <div><label style={MODAL_LBL}>Prix défaut (Ar)</label><input type="number" value={form.prix_defaut} onChange={e=>setForm(f=>({...f,prix_defaut:+e.target.value}))} style={MODAL_SI}/></div>
          <div style={{gridColumn:"1/-1"}}><label style={MODAL_LBL}>Mensualité défaut (Ar)</label><input type="number" value={form.mensualite_defaut} onChange={e=>setForm(f=>({...f,mensualite_defaut:+e.target.value}))} style={MODAL_SI}/></div>
        </div>
        <div style={{marginBottom:16}}><label style={MODAL_LBL}>Couleur</label><div style={{display:"flex",gap:8,marginTop:4}}>{COLS.map(c=><div key={c} onClick={()=>setForm(f=>({...f,couleur:c}))} style={{width:26,height:26,borderRadius:6,background:c,cursor:"pointer",border:`3px solid ${form.couleur===c?"#111":"transparent"}`,boxShadow:form.couleur===c?"0 0 0 2px #fff inset":""}}/>)}</div></div>
        <div style={{display:"flex",gap:8}}>
          <Btn onClick={onClose} variant="ghost" style={{flex:1}}>Annuler</Btn>
          <Btn onClick={()=>{if(form.nom&&form.sigle){form.emoji="·";onSave(form);}}} variant="solid" color={DS.accent} disabled={!form.nom||!form.sigle} style={{flex:2}}>Créer le site</Btn>
        </div>
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// ADMINISTRATION
// ═══════════════════════════════════════════════════════════════════
function AdminView({ users, session, onCreateUser, onUpdateUser, onDeleteUser }) {
  const [modal,setModal]=useState(null),[confirmDel,setConfirmDel]=useState(null);
  const ul=Object.values(users);
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
        <div><h2 style={{margin:0,fontSize:16,fontWeight:700,color:DS.text}}>Gestion des accès</h2><div style={{fontSize:11,color:DS.text4,marginTop:3}}>{ul.length} compte{ul.length>1?"s":""}</div></div>
        <Btn onClick={()=>setModal({type:"create"})} variant="solid" color={DS.accent}>+ Créer un accès</Btn>
      </div>
      <Card padding="0" style={{marginBottom:14}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr style={{background:"#f9fafb",borderBottom:`1px solid ${DS.border}`}}>{["Adresse email","Nom","Rôle","Statut","Actions"].map(h=><th key={h} style={{padding:"9px 14px",textAlign:"left",fontSize:10,fontWeight:600,color:DS.text4,textTransform:"uppercase",letterSpacing:"0.6px"}}>{h}</th>)}</tr></thead>
          <tbody>
            {ul.map((u,i)=>(
              <tr key={u.email} style={{borderBottom:`1px solid ${DS.border}`,background:i%2===0?"#fff":"#fafafa"}}>
                <td style={{padding:"11px 14px",fontSize:12,color:DS.blue}}>{u.email}</td>
                <td style={{padding:"11px 14px",fontSize:12,color:DS.text,fontWeight:500}}>{u.nom}</td>
                <td style={{padding:"11px 14px"}}><span style={{background:u.role==="admin"?DS.accentLt:DS.blueBg,color:u.role==="admin"?DS.accent:DS.blue,border:`1px solid ${u.role==="admin"?DS.accentMd:DS.blueBd}`,fontSize:10,padding:"2px 8px",borderRadius:4,fontWeight:600}}>{u.role==="admin"?"Administrateur":"Utilisateur"}</span></td>
                <td style={{padding:"11px 14px"}}><span style={{background:u.actif?DS.greenBg:DS.redBg,color:u.actif?DS.green:DS.red,border:`1px solid ${u.actif?DS.greenBd:DS.redBd}`,fontSize:10,padding:"2px 8px",borderRadius:4,fontWeight:600}}>{u.actif?"Actif":"Inactif"}</span></td>
                <td style={{padding:"11px 14px"}}>
                  <div style={{display:"flex",gap:6}}>
                    <Btn onClick={()=>setModal({type:"edit",user:u})} variant="ghost" size="sm">Modifier</Btn>
                    {u.email!==session.email&&<Btn onClick={()=>onUpdateUser(u.email,{actif:!u.actif})} variant={u.actif?"danger":"outline"} size="sm" color={u.actif?DS.red:DS.green}>{u.actif?"Désactiver":"Activer"}</Btn>}
                    {u.email!==session.email&&<Btn onClick={()=>setConfirmDel(u)} variant="danger" size="sm">Supprimer</Btn>}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {[{r:"Administrateur",d:"Accès complet à toutes les fonctionnalités, gestion des comptes utilisateurs, création et modification de sites.",c:DS.accent},{r:"Utilisateur",d:"Saisie de paiements, consultation des données, modification des dossiers acquéreurs, export Excel.",c:DS.blue}].map(({r,d,c})=>(
          <Card key={r}><div style={{fontSize:12,fontWeight:600,color:c,marginBottom:5}}>{r}</div><div style={{fontSize:12,color:DS.text3,lineHeight:1.6}}>{d}</div></Card>
        ))}
      </div>
      {modal?.type==="create"&&<ModalUser mode="create" onSave={d=>onCreateUser(d)} onClose={()=>setModal(null)}/>}
      {modal?.type==="edit"&&<ModalUser mode="edit" initial={modal.user} onSave={d=>onUpdateUser(modal.user.email,d)} onClose={()=>setModal(null)}/>}
      {confirmDel&&(
        <div style={{position:"fixed",inset:0,background:"rgba(17,24,39,0.4)",backdropFilter:"blur(4px)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}} onMouseDown={e=>{if(e.target===e.currentTarget)setConfirmDel(null);}}>
          <Card style={{width:380,boxShadow:DS.shadowLg}} onMouseDown={e=>e.stopPropagation()}>
            <div style={{fontSize:15,fontWeight:700,color:DS.text,marginBottom:6}}>Supprimer le compte</div>
            <div style={{fontSize:13,color:DS.text3,marginBottom:18,lineHeight:1.6}}>Confirmez-vous la suppression de <strong style={{color:DS.text}}>{confirmDel.nom}</strong> ({confirmDel.email}) ?</div>
            <div style={{display:"flex",gap:8}}><Btn onClick={()=>setConfirmDel(null)} variant="ghost" style={{flex:1}}>Annuler</Btn><Btn onClick={()=>{onDeleteUser(confirmDel.email);setConfirmDel(null);}} variant="solid" color={DS.red} style={{flex:1}}>Supprimer</Btn></div>
          </Card>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MODAL UTILISATEUR
// ═══════════════════════════════════════════════════════════════════
function ModalUser({ mode, initial, onSave, onClose }) {
  const isCreate=mode==="create";
  const [form,setForm]=useState({email:initial?.email||"",nom:initial?.nom||"",password:"",role:initial?.role||"user"});
  const [errMsg,setErrMsg]=useState("");
  function save(){
    setErrMsg("");
    if(!form.email.trim()||!form.nom.trim()){setErrMsg("Email et nom obligatoires.");return;}
    if(isCreate&&!form.password.trim()){setErrMsg("Mot de passe obligatoire.");return;}
    const p={email:form.email.toLowerCase().trim(),nom:form.nom.trim(),role:form.role};
    if(form.password.trim()) p.password=form.password.trim();
    const r=onSave(p);
    if(r&&r.ok===false){setErrMsg(r.msg||"Erreur.");return;}
    onClose();
  }
  const ok=form.email.trim()&&form.nom.trim()&&(isCreate?form.password.trim():true);
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(17,24,39,0.4)",backdropFilter:"blur(4px)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}} onMouseDown={e=>{if(e.target===e.currentTarget)onClose();}}>
      <Card style={{width:420,boxShadow:DS.shadowLg}} onMouseDown={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
          <div><div style={{fontSize:15,fontWeight:700,color:DS.text}}>{isCreate?"Créer un accès":"Modifier le compte"}</div>{!isCreate&&<div style={{fontSize:11,color:DS.text4,marginTop:2}}>{initial?.email}</div>}</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:DS.text4,cursor:"pointer",fontSize:22,lineHeight:1}}>×</button>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:11}}>
          <div><label style={MODAL_LBL}>Adresse email *</label><input type="email" value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))} disabled={!isCreate} placeholder="prenom.nom@exemple.mg" style={{...MODAL_SI,opacity:isCreate?1:0.55,cursor:isCreate?"text":"not-allowed"}}/>{!isCreate&&<div style={{fontSize:10,color:DS.text4,marginTop:3}}>L'email ne peut pas être modifié</div>}</div>
          <div><label style={MODAL_LBL}>Nom complet *</label><input value={form.nom} onChange={e=>setForm(f=>({...f,nom:e.target.value}))} placeholder="Prénom Nom" style={MODAL_SI}/></div>
          <div><label style={MODAL_LBL}>{isCreate?"Mot de passe *":"Nouveau mot de passe (vide = inchangé)"}</label><input type="text" value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))} placeholder={isCreate?"Définir le mot de passe":"Laisser vide pour conserver l'actuel"} style={MODAL_SI}/></div>
          <div><label style={MODAL_LBL}>Rôle</label><select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={MODAL_SI}><option value="user">Utilisateur — saisie, consultation, export</option><option value="admin">Administrateur — accès complet + gestion des accès</option></select></div>
        </div>
        {errMsg&&<div style={{marginTop:10,background:DS.redBg,border:`1px solid ${DS.redBd}`,borderRadius:6,padding:"8px 12px",color:DS.red,fontSize:12}}>{errMsg}</div>}
        <div style={{display:"flex",gap:8,marginTop:18}}>
          <Btn onClick={onClose} variant="ghost" style={{flex:1}}>Annuler</Btn>
          <Btn onClick={save} variant="solid" color={DS.accent} disabled={!ok} style={{flex:2}}>{isCreate?"Créer le compte":"Enregistrer"}</Btn>
        </div>
      </Card>
    </div>
  );
}
