Attribute VB_Name = "mdlInternalDatabase"
Public BahayaVBS(10) As String
Public BahayaBAT(10) As String
Public BahayaHTT(3) As String

Public Sub LoadVirusDatabase()
On Error Resume Next

Dim dbFile As String
Dim sData As String
Dim VirStr(1 To 2) As String
Dim strArray(666) As String
Dim I As Long

strArray(1) = "WORM.BADASS;1085DD16"
strArray(2) = "WORM.LOVGATE.W;10C08823"
strArray(3) = "WORM.ADRENALINE;11278EE4"
strArray(4) = "WORM.RODAL.A;118701D3"
strArray(5) = "BDS.1060.AL;1187C59B"
strArray(6) = "TH.ANNOYER;1237C3C1"
strArray(7) = "WORM.DREW.A;123CF8E8"
strArray(8) = "WORM.VB.BA;12494DA6"
strArray(9) = "BDS.CLI.UN.6;12661494"
strArray(10) = "WORM.LUNOR.A;12C31132"
strArray(11) = "TH.AF.20;12FDC721"
strArray(12) = "TH.ADUT;133CEF62"
strArray(13) = "WORM.BRONTOK.H;134E43F8"
strArray(14) = "TH.BAGLET.A;15CA7471"
strArray(15) = "TH.SMALL.VB.77824;15DA8F43"
strArray(16) = "TH.BSOD;15F00C9B"
strArray(17) = "WORM.MYDOOM.AD;1611CDED"
strArray(18) = "WGEN.SKYPEWORM;16A69630"
strArray(19) = "WGEN.KIT.PWG;16F750FC"
strArray(20) = "TH.KUAP;173A44FE"
strArray(21) = "TH.CONTHOL;174D3F73"
strArray(22) = "TH.ASNIFF.033;17D5AF6E"
strArray(23) = "WORM.MYDOOM.AE;18382639"
strArray(24) = "WORM.VIKING.N;183889A6"
strArray(25) = "WORM.APPFLET.A;188442D9"
strArray(26) = "WORM.DARBY.G2;18B8865A"
strArray(27) = "WORM.VB.BI;197A227C"
strArray(28) = "WORM.CHICKEN;198A3AEE"
strArray(29) = "TH.ATENDO;19B041F5"
strArray(30) = "TH.VB.LT.1;19D8BFD2"
strArray(31) = "TH.COLMATCH;1B1C8FFC"
strArray(32) = "TH.EVILTOOL;1B464C53"
strArray(33) = "WORM.LEGEMIR;1B9181FA"
strArray(34) = "WORM.AILIS.A;1BCAE0BB"
strArray(35) = "WORM.GAGA.A;1BDE3A03"
strArray(36) = "WORM.AGENT.ADL;1C140D67"
strArray(37) = "TH.CLICKME;1CA6038"
strArray(38) = "TH.ADIETH;1CABA1CA"
strArray(39) = "WORM.LISTAS;1CBAE18C"
strArray(40) = "WORM.VB.BE;1CE189F0"
strArray(41) = "WORM.NAVUP;1CE6DFA5"
strArray(42) = "TH.APAGAR;1D107794"
strArray(43) = "WORM.BOLGI;1D16F0B2"
strArray(44) = "WORM.HATI.A.1;1D30EB74"
strArray(45) = "WORM.FUNNYTAV.A;1DB4C96C"
strArray(46) = "WORM.ANIMAN;1DC8B613"
strArray(47) = "WORM.NOPLE;1DE511A0"
strArray(48) = "WORM.PLESA.A;1DEBBDF7"
strArray(49) = "WORM.GIFT.B;1E082B1E"
strArray(50) = "WORM.VIKING.D;1E0E8D7F"
strArray(51) = "W32.PLUTOR.B;1E32B2E"
strArray(52) = "TH.ANSPY;1E9BAAC1"
strArray(53) = "W32.YOSA.A;1EEA9E7D"
strArray(54) = "WORM.VB.BW.8;1F318629"
strArray(55) = "TH.BEFAST;1FEE7C2"
strArray(56) = "WORM.SORIW;2036232B"
strArray(57) = "WORM.VB.B;207F63DC"
strArray(58) = "WORM.BRONTOK.E;20C2C462"
strArray(59) = "WORM.LOVGATE.F;20C88AE7"
strArray(60) = "TH.NONAME.A;20F0504D"
strArray(61) = "TH.ANONEBOMBER.35;21B6ECA1"
strArray(62) = "WORM.MYDOOM.AP;21E2266F"
strArray(63) = "WORM.VB.AU;21F06DD1"
strArray(64) = "TH.FOTOEXE.A;226CE133"
strArray(65) = "EXPLOIT.MS05-039.A;228EDE27"
strArray(66) = "TH.TSK.N;22A2D45F"
strArray(67) = "TH.MAV.39;2370A29A"
strArray(68) = "TH.BUNGA;23A3DBCF"
strArray(69) = "TH.AUTOACCEPTER;246C02E4"
strArray(70) = "W32.XORALA.A;24AFD290"
strArray(71) = "TH.AFFC;24B33093"
strArray(72) = "WORM.VB.A;24ED0E87"
strArray(73) = "TH.ICH.1;250C66ED"
strArray(74) = "BDS.VB.AYT;2609FE7E"
strArray(75) = "VGEN.SBVM;26B01564"
strArray(76) = "TH.AZAK;2771517C"
strArray(77) = "WORM.ANTINNY.AE;291BD0F7"
strArray(78) = "WORM.ANTINNY.AG;2965D47B"
strArray(79) = "WORM.YAHLOVER.A;2A753D43"
strArray(80) = "TH.CAHYNA;2AD6994E"
strArray(81) = "TH.DLDR.SMALL.DDS;2B2FC6B"
strArray(82) = "WORM.WUN;2CB78064"
strArray(83) = "TH.BLACKHOLE;2D5C2952"
strArray(84) = "TH.AGGREVATOR;2E8A8750"
strArray(85) = "WORM.FOBOT;2EB523F6"
strArray(86) = "WORM.VB.GIMA.A;2FA116C3"
strArray(87) = "VGEN.THCK-FP;3080FC37"
strArray(88) = "WORM.VB.AZ.1;3088C3B"
strArray(89) = "WORM.VB.BZ.1;31065EE3"
strArray(90) = "WGEN.VBS.INDRA;31427976"
strArray(91) = "WORM.ANTITES;315A5F02"
strArray(92) = "WORM.PESIN.D;316C2031"
strArray(93) = "TH.AGENT.FVL;31D23C78"
strArray(94) = "VGEN.THCK-TC;322255DC"
strArray(95) = "TH.ARMOR.1;32E0098E"
strArray(96) = "WORM.MYDOOM.AH;33F59289"
strArray(97) = "TH.COPUPER;348EACDF"
strArray(98) = "WORM.VB.BY.4;34C18DA0"
strArray(99) = "TH.VB.NGM;353C6898"
strArray(100) = "WORM.TESTWORM;35620D35"
strArray(101) = "TH.CHICONS;3621F990"
strArray(102) = "WORM.DOLLY;36BA9F02"
strArray(103) = "TH.CRYPT.J;370039DC"
strArray(104) = "WORM.ALTICE;378007DC"
strArray(105) = "TH.VB.II;37AE1CBD"
strArray(106) = "VGEN.SWLABS.MACRO;37B774D5"
strArray(107) = "TH.CRAZYCD;37F2C04D"
strArray(108) = "WORM.DELF.AFK;39FD1B5D"
strArray(109) = "WORM.VB.H;3A372755"
strArray(110) = "WORM.STEFAN.A;3A801AAA"
strArray(111) = "TH.DROP.SHELLREQ.2;3AC0C027"
strArray(112) = "TH.COLORBUG;3AC505A7"
strArray(113) = "WORM.SHOLTA;3B4A9853"
strArray(114) = "WORM.VIKING.R;3BB79613"
strArray(115) = "TH.AGENT.ABT.3;3C6A366A"
strArray(116) = "TH.ANAME;3D561F06"
strArray(117) = "TH.DELF.KG.1;3D73F545"
strArray(118) = "WORM.DELF.CD;3D73F545"
strArray(119) = "TH.COLORER;3DB0B107"
strArray(120) = "WORM.TOUSH;3DDD7447"
strArray(121) = "TH.CRYPT.C;3E424AC5"
strArray(122) = "WORM.ANEL;3E4A78AC"
strArray(123) = "WORM.WILAB.A;3E509437"
strArray(124) = "WORM.ANTINNY.AD;3E93BBF0"
strArray(125) = "WORM.RBOT.210994;3EBF42AC"
strArray(126) = "TH.CRASH;3ED29B7B"
strArray(127) = "TH.AGENT.20480;3EE0179D"
strArray(128) = "TH.BANDEJA;3F0FB13"
strArray(129) = "WORM.SOHANNAD.NAK;3F18C707"
strArray(130) = "WORM.VB.EJ;3F42998A"
strArray(131) = "TH.JOINER.CM;3F980FDF"
strArray(132) = "WORM.VB.BS.13;40F5D3E2"
strArray(133) = "WORM.DAMEWAR;41EED650"
strArray(134) = "TH.VDX;4228AA77"
strArray(135) = "WORM.VB.BS.12;4249FB5A"
strArray(136) = "TH.CROWL;424B2F4C"
strArray(137) = "TH.AGENT.CGS;42A7A861"
strArray(138) = "WORM.VB.JM.14;436750A1"
strArray(139) = "WORM.VB.DB;4397FA0E"
strArray(140) = "TH.ABSTURZ;43A4BD35"
strArray(141) = "SPYWARE.MWS.381012;43BB4553"
strArray(142) = "TH.AUTHSTEALER;44299E9E"
strArray(143) = "WORM.BABUIN.A;443CD3B1"
strArray(144) = "TH.VB.HG.3;45868C4"
strArray(145) = "W32.VARWORM;4644BA47"
strArray(146) = "WORM.VB.BR;465A908A"
strArray(147) = "TH.ARCTICBOMB;465E5E3A"
strArray(148) = "WORM.NAUTICAL;46656368"
strArray(149) = "WORM.ABOTUS;46B1388F"
strArray(150) = "WORM.MYDOOM.AG;46E3B4AF"
strArray(151) = "TH.BARJAC;471A13BE"
strArray(152) = "WORM.VB.EA.1;480363F1"
strArray(153) = "WORM.DELF.AS.5;48E84785"
strArray(154) = "TH.PLUTO.A;490F0754"
strArray(155) = "TH.VB.BAF.1;4969517A"
strArray(156) = "TH.CRYPT.S;49CE2CD1"
strArray(157) = "WORM.WORMIC;49EE0B65"
strArray(158) = "TH.CORREO;4A06F60E"
strArray(159) = "WORM.BREAKER.A;4A8CADF8"
strArray(160) = "TH.BLINDER;4B732E6A"
strArray(161) = "TH.CARGAO.B;4BAE88E3"
strArray(162) = "TH.VB.AUS;4BBC2833"
strArray(163) = "TH.CSTAR.2;4C30C3BB"
strArray(164) = "WORM.VB.NKA;4C4C9F1D"
strArray(165) = "TH.CACOGEN;4D494E50"
strArray(166) = "WORM.VIKING.M;4D7FC452"
strArray(167) = "WORM.VB.L;4E19226E"
strArray(168) = "WORM.ALANIS;4E5529E7"
strArray(169) = "BDS.RABBIT.A;4EB44C35"
strArray(170) = "TH.VIRTL;4F5F04AA"
strArray(171) = "WORM.PLUTO.A.2;4FDCA109"
strArray(172) = "TH.BLEST;4FDFB052"
strArray(173) = "TH.LOP.KH;51FFAD3D"
strArray(174) = "WORM.BRONTOK.A;532F744C"
strArray(175) = "TH.BKCLIENT;5373E6DA"
strArray(176) = "VGEN.VB.R;53B8FA97"
strArray(177) = "WORM.VIKING.J;54E0C6B9"
strArray(178) = "WORM.SMALL.K.2;55EE21E6"
strArray(179) = "WORM.PASSMAIL.10;5619BD49"
strArray(180) = "TH.VB.PENGUIN.A;56A87DE2"
strArray(181) = "WORM.IRCGAME.A;56DC609"
strArray(182) = "WORM.VB.X;57D08894"
strArray(183) = "TH.STARBIT;586ACDCE"
strArray(184) = "TH.AIDLOT;58D8963"
strArray(185) = "TH.CRYPT.K;590462AE"
strArray(186) = "TH.CDDIE;597E9E9B"
strArray(187) = "VGEN.NUKE.GV;5A4A8EE"
strArray(188) = "TH.VB.AGB.2;5A56C72A"
strArray(189) = "MGEN.MMVG;5A7B3D0C"
strArray(190) = "TH.DROP.AQ.181252.B;5B0A30EB"
strArray(191) = "WORM.LIOTEN;5B74497C"
strArray(192) = "TH.KILLAV.HH.2;5BF92AC7"
strArray(193) = "TH.AIMBOT;5D84A99C"
strArray(194) = "TH.AIMBOT;5D84A99C"
strArray(195) = "TH.UNCODE;5DAB55B4"
strArray(196) = "WORM.VB.BAYA.1;5E677B89"
strArray(197) = "WORM.TOMRE.A;5E80CCB"
strArray(198) = "WORM.AMUS.A;5EB4EBFC"
strArray(199) = "WORM.BRONTOK.Q;5F628202"
strArray(200) = "TH.JANGKUROZ;60DCDB91"
strArray(201) = "WORM.MEXER;60E35C68"
strArray(202) = "TH.BELOCKER;61710251"
strArray(203) = "WORM.QAZ;62374F98"
strArray(204) = "WORM.LASTAS;6250B256"
strArray(205) = "WORM.STANBOX;6280B1D6"
strArray(206) = "TH.INDOVIRUS;6282A944"
strArray(207) = "TH.BAZOFIA;62D39F50"
strArray(208) = "TH.VB.BG.2;6316C432"
strArray(209) = "WORM.LENOVA;637C8606"
strArray(210) = "WORM.DECOY.A;64A2E8B5"
strArray(211) = "WORM.GROGOTIX.A;64EB1F6C"
strArray(212) = "TH.BINGO;65EFF01D"
strArray(213) = "W32.VB.AZ;65F74A1"
strArray(214) = "WORM.DELFER;66138018"
strArray(215) = "WORM.VB.FV;662235A3"
strArray(216) = "WORM.DAUR;668D8401"
strArray(217) = "TH.CRACK;670BCFD6"
strArray(218) = "TH.CHANTAL;67678DD0"
strArray(219) = "TH.BACTERIO61;677FA7F8"
strArray(220) = "WORM.ARCHMIME;681A614B"
strArray(221) = "WORM.VIKING.AD;686E6585"
strArray(222) = "WORM.ASID.A;68B232DF"
strArray(223) = "WORM.WINDAUS.B;68FCFB63"
strArray(224) = "TH.DLDR.STRATION.C;694DDCBC"
strArray(225) = "WORM.BRONTOK.I;69501808"
strArray(226) = "WORM.LEFROS.A;69967986"
strArray(227) = "WORM.POEBOT.81408.A;6A8A365B"
strArray(228) = "TH.DIEMODEM;6B03D141"
strArray(229) = "WORM.APLER;6B040A6D"
strArray(230) = "TH.WZ.A.1;6B940E31"
strArray(231) = "WORM.LITE.A;6C6CC490"
strArray(232) = "WORM.SLUTER;6C9FA287"
strArray(233) = "WORM.ANTINNY.M;6CA5BDFF"
strArray(234) = "TH.CRYPT.A;6CA64BE8"
strArray(235) = "VGEN.TROG.15;6D99383D"
strArray(236) = "TH.REQDIS.36864;6E08DC59"
strArray(237) = "W32.BLUR.RT;6E487929"
strArray(238) = "WORM.VB.BU.7;6E8B9B0A"
strArray(239) = "WORM.IMPONEX;6EB3DFFD"
strArray(240) = "WORM.VB.BD;7068CB15"
strArray(241) = "TH.CDGLUCK;71264BB6"
strArray(242) = "WORM.SMELLES;713B5FAE"
strArray(243) = "WORM.WINDAUS.C;720D5C15"
strArray(244) = "WORM.SHANSAI.A;7250970E"
strArray(245) = "WORM.LIMAR;72AB5F68"
strArray(246) = "WORM.VIKING.A;72BD5F6A"
strArray(247) = "WORM.FLUIKAN.D.9;732B09F5"
strArray(248) = "TH.C-KILLER;7390E60A"
strArray(249) = "WORM.SCORVAN;73B6DEC2"
strArray(250) = "TH.CRYPT.R;744BB7DF"
strArray(251) = "TH.CUKUX;75376914"
strArray(252) = "TH.CHERNICH;7572AE07"
strArray(253) = "WORM.AGENT.D;75A8CD22"
strArray(254) = "TH.VB.AGENT.E.2;75B172F6"
strArray(255) = "WORM.TOGUIVI.A;75ED9635"
strArray(256) = "WORM.VIKING.E;77450A72"
strArray(257) = "TH.COMPAIN;775552FC"
strArray(258) = "WORM.DENIT;775D70CC"
strArray(259) = "TH.HLLP.VB.J;7782CA4"
strArray(260) = "TH.ARMAGEDDON;77EE47FF"
strArray(261) = "TH.BUIZIT;785228DC"
strArray(262) = "TH.BROTHER;7858AE8"
strArray(263) = "TH.CUREICQ;793A50AB"
strArray(264) = "TH.CRYPT.F;79AEFC06"
strArray(265) = "WORM.VB.D;7A32D5D9"
strArray(266) = "TH.CHEAP.B;7B6CABEE"
strArray(267) = "VGEN.KIT.CM.N;7BAFBE17"
strArray(268) = "TH.BESYSAD.A;7BBA9EBE"
strArray(269) = "WORM.ANTINNY.AF;7C019038"
strArray(270) = "TH.JSEB;7C74312C"
strArray(271) = "TH.TL.G;7C74312C"
strArray(272) = "TH.DROP.HI.467976.B;7C84A756"
strArray(273) = "TH.FIRMALEX;7D12F960"
strArray(274) = "W32.SETINS.V;7D3CAF51"
strArray(275) = "WORM.DELINF;7D694B0E"
strArray(276) = "TH.BERTZ;7D8159D5"
strArray(277) = "TH.CRYPT.T;7DA1DE1"
strArray(278) = "WORM.HOMET.A;7DE417CF"
strArray(279) = "WORM.ANTINNY.AH;7E1CE8E0"
strArray(280) = "WORM.ULTIMAX.B;7FE5E9D"
strArray(281) = "TH.CRUI.F;800BB857"
strArray(282) = "TH.VB.ANR;802B6599"
strArray(283) = "TH.NETGHOST;802C1214"
strArray(284) = "TH.AUTORUN.BW;8164D5AD"
strArray(285) = "WORM.TRAXQ.B;822CF830"
strArray(286) = "W32.DELF.AC;8277B1DA"
strArray(287) = "WORM.VB.DH.5;8278D88D"
strArray(288) = "WORM.SOHANAD.AE.1;828078D0"
strArray(289) = "TH.WO.4;83509DA0"
strArray(290) = "WORM.TULU;83745DEC"
strArray(291) = "WORM.ZWQQ.A;8416F936"
strArray(292) = "TH.CRYPT.O;849E19C2"
strArray(293) = "TH.FAKEFORMAT;853DD1BF"
strArray(294) = "WORM.PENDEX;859E348A"
strArray(295) = "DR.BRONTOK.J.1;85B1A381"
strArray(296) = "WORM.VB.CM;85E19B67"
strArray(297) = "TH.AGENT.89041;86DD9E29"
strArray(298) = "WORM.HOBGOB.A;86EBE2D6"
strArray(299) = "WORM.ANTINNY.AO;870AAD0B"
strArray(300) = "TH.DCOM.AD.2;870AF38C"
strArray(301) = "WORM.APLORE;887C36DE"
strArray(302) = "TH.BOMBAT;8B1343FC"
strArray(303) = "WORM.PARVED.B;8B80F213"
strArray(304) = "WORM.SMALL.K.1;8BBBD868"
strArray(305) = "WORM.VB.AN;8BF7CB"
strArray(306) = "TH.COLDFUSION;8C156CE6"
strArray(307) = "TH.DEL.POM;8C353568"
strArray(308) = "WORM.MAGCALL;8C53FED0"
strArray(309) = "WORM.TAK.A;8D14BE7F"
strArray(310) = "TH.CHRIST;8D3E6F58"
strArray(311) = "TH.CUHMAP;8D3E7C6F"
strArray(312) = "WORM.MIRCUP;8D5113C1"
strArray(313) = "TH.AGENT.DBH.1;8DE95E7B"
strArray(314) = "WORM.ALCAUL;8E0A0B40"
strArray(315) = "TH.COSTARO;8E3785C9"
strArray(316) = "WORM.VB.AY.2;8E444A69"
strArray(317) = "WORM.VB.CG;8EC7F7F2"
strArray(318) = "WORM.SETEADA;8F275CF2"
strArray(319) = "SPYWARE.MWS.57344;90564BC5"
strArray(320) = "TH.MSGBOX.A;90594177"
strArray(321) = "TH.CRYPT.E;9105D937"
strArray(322) = "TH.AMBER;9109031B"
strArray(323) = "TH.ANALOX;91699B1E"
strArray(324) = "TH.BUM;918BC58F"
strArray(325) = "TH.FAKEBSOD;919EEE2A"
strArray(326) = "TH.DLDR.DEL;91EC303C"
strArray(327) = "WORM.RJUMP.A;923F2F0A"
strArray(328) = "WORM.MANEX;92666300"
strArray(329) = "TH.VB.AKZ;93B711F8"
strArray(330) = "TH.ADDUSER.N;94AFD05"
strArray(331) = "VGEN.IVSC;9532F684"
strArray(332) = "TH.CARDS;95FC8DD1"
strArray(333) = "TH.BUNLK;960C0A7B"
strArray(334) = "TH.DRONE.29;96136F90"
strArray(335) = "WORM.DENISBEE;96871D90"
strArray(336) = "WORM.PASSMA;96E7DD55"
strArray(337) = "WORM.ASTIX;97328E44"
strArray(338) = "TH.ADGOBLIN;97621DC8"
strArray(339) = "BDS.FLYVB;9802DA60"
strArray(340) = "VGEN.NGVCK.145;9804D951"
strArray(341) = "WORM.AIDID;9887031C"
strArray(342) = "WORM.CURGIRL.IRC;9970E360"
strArray(343) = "WORM.VB.AA;99B46DC3"
strArray(344) = "WORM.HUAYU.A;9A585C55"
strArray(345) = "WORM.VIKING.BG;9B28C04F"
strArray(346) = "WORM.CRYBOT.A;9C3AEAD"
strArray(347) = "WORM.WINDAUS.D;9CD181EA"
strArray(348) = "WORM.FLEMING;9CD6ECA9"
strArray(349) = "TH.CRYPT.Q;9CDA7951"
strArray(350) = "TH.AKUAN;9CF25275"
strArray(351) = "WORM.BRONTOK.N;9D1A8627"
strArray(352) = "WORM.VIKING.K;9E70C321"
strArray(353) = "WORM.BRONTOK.J;9EA65143"
strArray(354) = "TH.CRYPT.B;9EBA61BE"
strArray(355) = "WORM.ASSARM;9FD3BA85"
strArray(356) = "W32.VB.Q;9FD4B16A"
strArray(357) = "TH.ALERTA;A0230261"
strArray(358) = "WORM.KANGEN.A;A04B4E54"
strArray(359) = "TH.ALARM.A;A11CF620"
strArray(360) = "WORM.NUF;A18A95A3"
strArray(361) = "TH.WASS;A1AB2C4E"
strArray(362) = "WORM.MYDOOM.AN;A2453C5F"
strArray(363) = "WORM.JBK.A.3;A26BE35A"
strArray(364) = "WORM.FELIX;A2E39F3"
strArray(365) = "WORM.VB.AT;A36B1A56"
strArray(366) = "TH.PIRTES.A;A382C673"
strArray(367) = "WORM.BRONTOK.F;A443459E"
strArray(368) = "WORM.ANTINNY.AK;A46CF485"
strArray(369) = "TH.CSKEY;A49512BD"
strArray(370) = "WORM.ANTINNY.AJ;A4CA0230"
strArray(371) = "WORM.VIKING.H;A5310773"
strArray(372) = "TH.NETDEMON;A53C45C3"
strArray(373) = "TH.PSW.VB.JP.11;A5F73673"
strArray(374) = "TH.VB.I;A60CD1D"
strArray(375) = "WORM.BRONTOK.B;A6D132AD"
strArray(376) = "WORM.AGENT.B.18;A71ABBA4"
strArray(377) = "TH.CONLOCK;A791B799"
strArray(378) = "WORM.RAHAK.A;A878E568"
strArray(379) = "WORM.NAXE;A925EF87"
strArray(380) = "TH.AJIM;A97FA6E5"
strArray(381) = "WORM.DESIRE;A9B6BD9F"
strArray(382) = "TH.BOOMER;AA2E516F"
strArray(383) = "TH.VB.SICUFFIT;AA3A9D9F"
strArray(384) = "WORM.PLURED.B;AA79D928"
strArray(385) = "WORM.BRONTOK.416284;AAD57DDE"
strArray(386) = "WORM.MYDOOM.AQ;AAECA8D"
strArray(387) = "WORM.BRONTOK.X;AB48FBA1"
strArray(388) = "TH.BDH.1.A;ABB363C6"
strArray(389) = "WORM.BRONTOK.C;ABE7AD1B"
strArray(390) = "TH.CURHU;AC7E2AED"
strArray(391) = "TH.ANUBIS.110;AD75313B"
strArray(392) = "WORM.VB.CP;AD9C8D8C"
strArray(393) = "WORM.DUELLA.A;ADBC55BB"
strArray(394) = "TH.BLUEBOY;AE66DB24"
strArray(395) = "WORM.SALITY.C.2;AF313E4F"
strArray(396) = "TH.NETBUS.311;AF519870"
strArray(397) = "WORM.BILAY.A;AF831609"
strArray(398) = "WORM.KULLAN;AFA38130"
strArray(399) = "TH.VB.EMU;B01FA309"
strArray(400) = "TH.LOOPS.DLL.A.1;B10A1BC0"
strArray(401) = "TH.SPY.VB.TQ;B1C9420A"
strArray(402) = "WORM.VIKING.B;B3012D33"
strArray(403) = "W32.NIMDA;B361AE26"
strArray(404) = "W32.VB.BY;B437763A"
strArray(405) = "TH.CAMKING;B47DC50"
strArray(406) = "TH.VB.ATW;B4C1CE6A"
strArray(407) = "WORM.VB.CB;B4DF7A01"
strArray(408) = "WORM.ANTINNY.L;B53D5A87"
strArray(409) = "WORM.BUGUS.A;B5AA1699"
strArray(410) = "WGEN.KIT.VBS.2.A;B5DC73C6"
strArray(411) = "TH.CHICO;B6627DC7"
strArray(412) = "WORM.VIKING.AA;B6DAF89"
strArray(413) = "WORM.ANPIR.A;B8338D9D"
strArray(414) = "VGEN.CWG.A;B843D7C"
strArray(415) = "TH.ICABDI.B;BA45DB5F"
strArray(416) = "TH.CLOZ.A;BB60BDFA"
strArray(417) = "TH.BSON;BB737C87"
strArray(418) = "WORM.BRONTOK.L;BBE0C6D2"
strArray(419) = "WORM.VB.S;BCC4CD45"
strArray(420) = "TH.CDARGEN;BD18DA53"
strArray(421) = "TH.BOA;BD19234C"
strArray(422) = "WORM.GAVIR;BD1A5C67"
strArray(423) = "TH.CHOO.B;BD279D1E"
strArray(424) = "TH.ALFOOL;BD6FA910"
strArray(425) = "WORM.ANTINNY.K;BD879B64"
strArray(426) = "WORM.ANTINNY.V;BD879B64"
strArray(427) = "TH.VB.W;BDCF3DE9"
strArray(428) = "WORM.ZAURGA.A;BE472B60"
strArray(429) = "TH.CABLEBOOST;BEBFDE43"
strArray(430) = "TH.VB.SMALL.88442;BECE376"
strArray(431) = "WORM.LSAN;BED529B4"
strArray(432) = "TH.EZ31;BFC03328"
strArray(433) = "TH.ANGRIFF;BFE66171"
strArray(434) = "WORM.BRONTOK.W.1;C013B812"
strArray(435) = "WORM.KOBER;C06CDA6B"
strArray(436) = "TH.CPUHOG.10;C0968591"
strArray(437) = "WORM.TARK.A;C13193F2"
strArray(438) = "TH.CRYPT.D;C16AB34E"
strArray(439) = "WORM.VB.C;C1E61B18"
strArray(440) = "VGEN.BEEBS;C23E56B6"
strArray(441) = "WORM.BRONTOK.O;C25E70EA"
strArray(442) = "W32.XORALA.B;C2C3361"
strArray(443) = "WORM.TORUN;C340632F"
strArray(444) = "WORM.3DSTARS;C408CEA4"
strArray(445) = "TH.COKEGIFT;C4800644"
strArray(446) = "TH.ABADDON;C5A3AD7C"
strArray(447) = "TH.VB.LG;C5F28070"
strArray(448) = "TH.AUTOIT.AX.1;C69C753C"
strArray(449) = "WORM.WARPIGS.A;C6F1018A"
strArray(450) = "WORM.CZ.14.A;C803B8AB"
strArray(451) = "W32.RELIC.B;C813C4CA"
strArray(452) = "WORM.MYDOOM.AL;C8DF0248"
strArray(453) = "TH.ADDER;C8E38C1C"
strArray(454) = "TH.DESK.1;C91C269A"
strArray(455) = "TH.DELF;C97A59B5"
strArray(456) = "TH.BAYAN;C99C7D7"
strArray(457) = "WORM.KANGEN.B;C9BC8192"
strArray(458) = "TH.VIXEN.A;C9E7E9CF"
strArray(459) = "TH.FLOODER.IM.VB.GC;CA3EC214"
strArray(460) = "TH.MATRIX;CBA968B1"
strArray(461) = "WORM.RIDNU.D;CC01720F"
strArray(462) = "VGEN.SW32OWG;CC11CD20"
strArray(463) = "VGEN.KIT.ANSIMAKE.B.2;CCBD23BE"
strArray(464) = "TH.CONIP;CCD81B82"
strArray(465) = "WORM.ANTINNY.C;CECD5FAC"
strArray(466) = "WORM.ANTINNY.X;CECD5FAC"
strArray(467) = "BDS.PERDOR.A;CF16B4B1"
strArray(468) = "WORM.DROF;CFD6B8E4"
strArray(469) = "WORM.ZIPPY;D08E4702"
strArray(470) = "TH.BEROK;D0AC9F10"
strArray(471) = "TH.ALMAEDA;D0EE57BB"
strArray(472) = "WORM.AGIST.A;D1135D69"
strArray(473) = "WORM.AVONER;D15BCB9C"
strArray(474) = "W32.PESIN.B;D15FC2C6"
strArray(475) = "TH.SCR.V;D1C356E6"
strArray(476) = "TH.CONFIGLOOP;D1D92EEF"
strArray(477) = "WORM.MYDOOM.AJ;D25BF85"
strArray(478) = "WORM.AREQUIPA.B;D2ADB338"
strArray(479) = "WORM.RAYS;D2C1731E"
strArray(480) = "WORM.BRONTOK.G;D2ED166F"
strArray(481) = "WORM.NEWFOLDER.A;D30064E3"
strArray(482) = "WORM.VB.E;D354E687"
strArray(483) = "TH.SPY.BAIDU;D4045169"
strArray(484) = "TH.BUTANO;D415F58C"
strArray(485) = "WORM.VIKING.AS;D4640F81"
strArray(486) = "TH.MAILBOMB.02;D5255EEB"
strArray(487) = "WORM.VITAN.A;D53AC6E6"
strArray(488) = "WORM.FLUIKAN.A.1;D5C2385F"
strArray(489) = "WORM.BRONTOK.P;D623748F"
strArray(490) = "VGEN.THCK-TBC;D676D561"
strArray(491) = "TH.CSTAR;D716B3CC"
strArray(492) = "WORM.NETSKY.AP;D76D1A96"
strArray(493) = "WORM.MOOZE;D83DC75"
strArray(494) = "WORM.ANTINNY.B;D90AC03E"
strArray(495) = "WORM.ANARCH;DA06A31C"
strArray(496) = "TH.BANCOS.PWW;DA586A25"
strArray(497) = "TH.CAPIRUF;DA87F37"
strArray(498) = "WORM.WARPIGS.B;DACC82AB"
strArray(499) = "TH.COMSN;DBB6EF69"
strArray(500) = "WORM.BRONTOK.K;DBF2E57F"
strArray(501) = "WORM.LADEX.A;DC357741"
strArray(502) = "WORM.VB.P;DC5002E8"
strArray(503) = "WORM.FOLDER.VB.A;DCF484D1"
strArray(504) = "TH.SPY.LOOPS.A;DD07C6AB"
strArray(505) = "WORM.NETBOT.A;DD61709E"
strArray(506) = "WORM.VB.X.3;DD832A41"
strArray(507) = "TH.BUTTONF;DE4971E8"
strArray(508) = "VGEN.ELIM.G;DE9AA5C7"
strArray(509) = "TH.AGENT.CEW;DEA7E363"
strArray(510) = "TH.EX.48;DF19F7BF"
strArray(511) = "TH.AVOID.MB;DFC32893"
strArray(512) = "TH.COOL;DFD7A856"
strArray(513) = "VGEN.G2;E1B940A2"
strArray(514) = "WORM.MYDOOM.AA;E242A69A"
strArray(515) = "WORM.BABYBEAR;E28C09B6"
strArray(516) = "TH.AGENT.XAD;E28DA6EA"
strArray(517) = "WORM.PAKOTA.C;E2A4B166"
strArray(518) = "WORM.ALPHX.B;E2DEC6D6"
strArray(519) = "WORM.VB.BU;E2E2B7C8"
strArray(520) = "TH.NORUTA.A;E3832847"
strArray(521) = "WORM.PERSER;E3960A71"
strArray(522) = "WORM.SYSDIL;E3CC4ED9"
strArray(523) = "WORM.TUTIAM.A;E41224FA"
strArray(524) = "TH.CLEANLOGS;E53206A6"
strArray(525) = "WORM.SILENTIUM.A;E5515A8D"
strArray(526) = "WORM.VIKING.Y;E57BD9A1"
strArray(527) = "VGEN.VCKIT;E5AFE382"
strArray(528) = "TH.BLOCCO;E6B07D15"
strArray(529) = "WORM.BRONTOK.D;E6CBFC25"
strArray(530) = "TH.XINXIN;E73FCD30"
strArray(531) = "TH.PALODNI.A;E76C6B4D"
strArray(532) = "TH.DLDR.AGENTT.II.5.B;E7A05EE9"
strArray(533) = "WORM.ONVER;E7B28268"
strArray(534) = "DR.ZLOB.N;E832EA7E"
strArray(535) = "VGEN.FCONV;E88249E1"
strArray(536) = "VGEN.SATANIC.A;E8B5192E"
strArray(537) = "WORM.WENPER.B;E8EEBB26"
strArray(538) = "WORM.EVOLMI;E96152BD"
strArray(539) = "WORM.BRONTOK.S.1;E997FC89"
strArray(540) = "WORM.BRONTOK.V.1;EA35E932"
strArray(541) = "WORM.DREFIR.A;EAB2491A"
strArray(542) = "WORM.NOWIM;EAF072F8"
strArray(543) = "WORM.MYDOOM.AK;EB9B99CB"
strArray(544) = "TH.BLOODLUST;EBD91534"
strArray(545) = "TH.FLOOD.AVRIL;EC1EFA5A"
strArray(546) = "WGEN.VBS.PSWVG;EC3A0451"
strArray(547) = "TH.VB.ABY.5.A;EC7EF1B1"
strArray(548) = "TH.CVIH;ECB202B3"
strArray(549) = "WORM.DERIUM.A;ED1F2F21"
strArray(550) = "TH.VB.AEI;ED92EBE9"
strArray(551) = "WORM.VB.N;ED94523"
strArray(552) = "TH.VB.BG;ED970D47"
strArray(553) = "TH.DRLDL.FAN.A;EDE7084E"
strArray(554) = "WORM.WILAB.D;EDE84BD4"
strArray(555) = "TH.KILLAV.HH;EE2013AA"
strArray(556) = "TH.CENTERO;EE2CA8F0"
strArray(557) = "WORM.DUPATE.4180;EED2DAB4"
strArray(558) = "TH.VB.AVD;EFA9CC3B"
strArray(559) = "WORM.BUMERANG;F0DA4B10"
strArray(560) = "WORM.VB.MOON.BN;F0E2F35D"
strArray(561) = "TH.COVERT;F13411B"
strArray(562) = "WORM.HAI;F1BC96A7"
strArray(563) = "WORM.NETOL;F1DF5F8C"
strArray(564) = "WORM.ANTINNY.H;F2178FDF"
strArray(565) = "WORM.AGENT.I;F235ED6B"
strArray(566) = "TH.CUKI;F23CF330"
strArray(567) = "W32.VB.AZ.2;F2EDA45C"
strArray(568) = "WORM.ANTINNY.I;F34DA1F5"
strArray(569) = "WORM.ANTINNY.R;F38A8463"
strArray(570) = "WORM.FLYING;F3EFABDC"
strArray(571) = "WORM.MYDOOM.AF;F4EFC64C"
strArray(572) = "WORM.VB.AV;F57BDA69"
strArray(573) = "WORM.ANTINNY.J;F5ACEC45"
strArray(574) = "TH.DELF.XP;F5BD0FD4"
strArray(575) = "WORM.ULTIMAX.A;F5C971AF"
strArray(576) = "WORM.DISCOBALL;F6422B9"
strArray(577) = "WORM.FOXMA;F651C540"
strArray(578) = "TH.BLACKBIRD;F6B01C81"
strArray(579) = "WORM.KILLAV.GR;F6F76C88"
strArray(580) = "WORM.LUNATIK;F702396F"
strArray(581) = "WORM.VB.M;F7598FC"
strArray(582) = "WORM.VIKING.G;F78CF22"
strArray(583) = "WORM.VB.CZ.5;F7ECBFAD"
strArray(584) = "WORM.VB.DR.12;F85E3831"
strArray(585) = "WORM.VB.CX;F8DAAED9"
strArray(586) = "TH.CHOKE;F9830037"
strArray(587) = "SPYWARE.AGENT.92160;F99AC66A"
strArray(588) = "WORM.VB.CK;F9DBF10B"
strArray(589) = "WORM.AZRAEL;FB6C6280"
strArray(590) = "TH.CAREM;FB7C29D9"
strArray(591) = "WORM.CISUM.A;FBC7E301"
strArray(592) = "WORM.SEESIX;FC622EAA"
strArray(593) = "TH.DELF.BVS.4;FDB41C6"
strArray(594) = "WORM.DOSIG.A;FE17AB74"
strArray(595) = "WORM.VIKING.I;FE3E45AC"
strArray(596) = "TH.COBRA;FE68F54A"
strArray(597) = "TH.SMALL.JB;FE7C3213"
strArray(598) = "WORM.OFFBOT.A;FE89272D"
strArray(599) = "WORM.MYDOOM.AI;FF016332"
strArray(600) = "TH.CLIDEM;FF925B78"
strArray(601) = "TH.XXCPP;FFAD1948"
strArray(602) = "WORM.BRONTOK.M;FFDC6CA8"
strArray(603) = "TH.DROP.LOOPS.A.1;5D6EE3C3"
strArray(604) = "TH.DROP.VB.DU.1;12BB706A"
strArray(605) = "TH.VB.NJO;2E6D645B"
strArray(606) = "TH.KAMERA.AMU.3;AFC4E167"
strArray(607) = "TH.KAMERA.AMU;5C6EB926"
strArray(608) = "W32.VB.CC;6B47B31E"
strArray(609) = "TH.VB.AOY.2;79487D34"
strArray(610) = "TH.VB.PG;C52E528"
strArray(611) = "TH.KAMERA.AMU.2;D0E391F"
strArray(612) = "WORM.WKYO86;13F301EA"
strArray(613) = "W32.STARKID;214FB8B"
strArray(614) = "VGEN.ANGSA;8DD3D043"
strArray(615) = "W32.CANABIS.1;61F5D444"
strArray(616) = "W32.CANABIS.2;F92A8D12"
strArray(617) = "W32.CANABIS.3;ED14C43C"
strArray(618) = "W32.OROCHIMARU;998288FC"
strArray(619) = "W32.PEND.BLANK;97276510"
strArray(620) = "W32.VB.SAMS;DBB571E8"
strArray(621) = "W32.VB.SWFDC;9FDECA34"
strArray(622) = "W32.XEROR;C5783BB6"
strArray(623) = "TH.GENERIC PWS.Y;90463CC0"
strArray(624) = "W32.BACALID;97276510"
strArray(625) = "WORM.SUSPECTED;F8D93D96"
strArray(626) = "WORM.SUSPECTED;F6032FB4"
strArray(627) = "BSC.KALONG;B0535741"
strArray(628) = "BSC.NIHLIT;21B16D19"
strArray(629) = "W32.AMBURADUL;A1B77387"
strArray(630) = "W32.NEW MALWARE;76096BDD"
strArray(631) = "W32.SALITY.AC;A92F8372"
strArray(632) = "W32.GENERIC.WORM!IRC;D1519D58"
strArray(633) = "W32.VIRUT.GEN.A;2CC8FB16"
strArray(634) = "W32.DX;B07071AB"
strArray(635) = "W32.ALMANAHE.C;8AC5DF32"
strArray(636) = "W32.GENERIC.A@MM;508B364F"
strArray(637) = "W32.MALWARE.A!ZIP;DA0825B7"
strArray(638) = "W32.CEP.WORM!33925D66;9C5FF009"
strArray(639) = "W32.MOONTOX-BRO.VARIANT;D7E50EC4"
strArray(640) = "W32.MOONTOX-BRO.C(DANGER);AB0E9785"
strArray(641) = "W32.DX;60A13250"
strArray(642) = "W32.CEKAR;17944E77"
strArray(643) = "W32.SALITY.AG;3366BD46"
strArray(644) = "W32.BACKDOOR-ASC.CFG;1F9F7225"
strArray(645) = "W32.DROPPER;3F980FDF"
strArray(646) = "W32.SPAM-MAILBOT;735C212C"
strArray(647) = "BSC.FERTP;432187F0"
strArray(648) = "BSC.BACKDOOR;C18A1AA9"
strArray(649) = "W32.DK;41FDC70A"
strArray(650) = "W32.VIRUT.GEN;3F2B36AE"
strArray(651) = "TH.QQROB;A0A49D42"
strArray(652) = "TH.AUTORUN;C4733C17"
strArray(653) = "W32.NARUTO;B2B7F64D"
strArray(654) = "TH.BLACK CIRCLE.KIT;28682F6C"
strArray(655) = "W32.SALITY.AO;EC6AC9B8"
strArray(656) = "W32.SALITY.AO;30996E3"
strArray(657) = "W32.SALITY.AO;BCD0006A"
strArray(658) = "W32.SALITY.AO;27AE19EA"
strArray(659) = "W32.VB.B;DEA7E363"
strArray(660) = "W32.EF;67D945E"
strArray(661) = "W32.SALITY.AO;5E7F25D9"
strArray(662) = "W32.SALITY.AO;3EEAF894"
strArray(663) = "W32.SORACI;F130176"
strArray(664) = "W32.BACKDOOR;BA7E1BEE"
strArray(665) = "W32.KESPO.A;DFE311E3"
strArray(666) = "W32.VCC2000.KIT;F0580E54"

    For I = 1 To UBound(strArray)
        VirStr(1) = Split(strArray(I), ";")(0)
        VirStr(2) = Split(strArray(I), ";")(1)
        VirusName.Add VirStr(1)
        VirusSign.Add VirStr(2)
    Next I
    
    MkDir nPath(App.path) & "Quarantine"
End Sub

Public Sub LoadBinaryIconCompare()

On Error Resume Next

Dim dbFile As String
Dim sData As String
Dim VirStr(1 To 2) As String
Dim strArray(619) As String
Dim I As Long

strArray(1) = "AKSIKA;20938B2"
strArray(2) = "KANGEN;19F4ED6"
strArray(3) = "APEL;133BE0B"
strArray(4) = "APEL;18EDEAE"
strArray(5) = "BRONTOK;1EF89C2"
strArray(6) = "ARMORA;1C915FF"
strArray(7) = "ASCRIBES;24563C4"
strArray(8) = "CODEX;1B2DB74"
strArray(9) = "BRONTOK;208EA72"
strArray(10) = "CYRAX;22A064D"
strArray(11) = "CYRAX;19B64EE"
strArray(12) = "DECOIL;1D4B7E1"
strArray(13) = "ROLOG;2087762"
strArray(14) = "EGO;29C7258"
strArray(15) = "FLUBURUNG;1B18705"
strArray(16) = "GELAS;1B5FCAB"
strArray(17) = "IMELDA;126D4CF"
strArray(18) = "IMELDA;1C58E5C"
strArray(19) = "IWING;15D7730"
strArray(20) = "JABLAY;1FB82B7"
strArray(21) = "KAMASUTRA;112763E"
strArray(22) = "LEENA;2165AF9"
strArray(23) = "MAZDA;25F46BE"
strArray(24) = "MYSONG;206556B"
strArray(25) = "NAHITAL;22A8D69"
strArray(26) = "NETSKY;19237F8"
strArray(27) = "RIYANI;15022B4"
strArray(28) = "NIMDA;1D8B4EB"
strArray(29) = "NUKEDEVIL;1DBC1EA"
strArray(30) = "PARAYRONTOK;2333F5D"
strArray(31) = "PETA;1F37C2F"
strArray(32) = "PLUTO;1C9CCA4"
strArray(33) = "PLUTO;1DFDFB4"
strArray(34) = "POLYFACE;1C1283E"
strArray(35) = "PROVISIONING;1F6598C"
strArray(36) = "RENOVA;27F4C1A"
strArray(37) = "STRATION;22F92E0"
strArray(38) = "TINUTUAN;191DBDC"
strArray(39) = "TSUNAMI;27BFE4A"
strArray(40) = "WUKILL;20E0907"

    For I = 1 To UBound(strArray)
        VirStr(1) = Split(strArray(I), ";")(0)
        VirStr(2) = Split(strArray(I), ";")(1)
        IconName.Add VirStr(1)
        IconSign.Add VirStr(2)
    Next I
    
End Sub

Public Sub InitScriptHeuristic()
On Error Resume Next

'VBS Script
    BahayaVBS(0) = "scripting.filesystemobject"
    BahayaVBS(1) = "wscript.shell"
    BahayaVBS(2) = "wscript.scriptfullname"
    BahayaVBS(3) = "regsetvalue"
    BahayaVBS(4) = "copyfile"
    BahayaVBS(5) = "exitwindowsex"
    BahayaVBS(6) = "persistmoniker=file:"
    BahayaVBS(7) = "runit"
    BahayaVBS(8) = "attachments.add"
    BahayaVBS(9) = "outlook.application"
    BahayaVBS(10) = "worm"
    
'BAT Script
    BahayaBAT(0) = "format "
    BahayaBAT(1) = "reg "
    BahayaBAT(2) = "%0"
    BahayaBAT(3) = "attrib "
    BahayaBAT(4) = "\run"
    BahayaBAT(5) = "hidden"
    BahayaBAT(6) = "disable"
    BahayaBAT(7) = "startup"
    BahayaBAT(8) = "NoFolderOptions"
    BahayaBAT(9) = "HideFileExt"
    BahayaBAT(10) = "tskill "

'HTT Script
    BahayaHTT(0) = "runexe"
    BahayaHTT(1) = "document.writeln(runexe)"
    BahayaHTT(2) = "object id=\""runit\"" width=0 height=0 type=\""application/x-oleobject\"
    BahayaHTT(3) = "codebase=\"
End Sub

Public Function HitDatabase()
    Dim I As Integer
    Dim vCount As Integer
    
    frmScanVirus.lstVirus.ListItems.Clear
    vCount = 0
    For I = 1 To VirusName.count
        vCount = vCount + 1
        frmScanVirus.lstVirus.ListItems.Add , , VirusName.Item(I), , 1
    Next I

    With frmScanVirus
        .lblSystem(0).Caption = ": " & CURRENT_BUILD
        .lblSystem(1).Caption = ": " & ENGINE_VERSION
        If .lstVirus.ListItems.count <> 0 Then
            .lblSystem(2).Caption = ": " & .lstVirus.ListItems.count & " " & "Viruses"
        Else
            .lblSystem(2).Caption = ": 0 Viruses"
        End If
    End With
    
End Function
