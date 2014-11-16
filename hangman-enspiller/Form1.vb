Public Class Form1

    Private spillmatrise() As Char 'Ordet som brukes i denne runden
    Private ordliste(468) As String 'Alle ordene som kan brukes i spillet

    Private spillordlengde As Integer 'Lengden på ordet som brukes i denne runden
    Private liv As Integer 'Antallet liv/streker på tegningen spilleren har til rådighet hver runde
    Private bokstaverIgjen As Integer 'Antallet bokstaver som gjenstår å finne denne runden
    Private spillAvLyd As Boolean 'Lagrer lydinnstilling

    'Prosedyre som laster inn ordlisten i matrisen
    Private Sub lastInnOrdliste()
        ordliste(0) = "BLEIE"
        ordliste(1) = "BART"
        ordliste(2) = "FJERNSYN"
        ordliste(3) = "BIL"
        ordliste(4) = "TELEFON"
        ordliste(5) = "ARKITEKT"
        ordliste(6) = "VINFLASKE"
        ordliste(7) = "LETTØL"
        ordliste(8) = "KALKUN"
        ordliste(9) = "HÅRBØRSTE"
        ordliste(10) = "BRUNKREM"
        ordliste(11) = "SMØRBUKK"
        ordliste(12) = "TASTATUR"
        ordliste(13) = "WINDOWS"
        ordliste(14) = "DATAMASKIN"
        ordliste(15) = "SKJERM"
        ordliste(16) = "BOK"
        ordliste(17) = "HULLMASKIN"
        ordliste(18) = "PENAL"
        ordliste(19) = "BLYANT"
        ordliste(20) = "GENSER"
        ordliste(21) = "TRENING"
        ordliste(22) = "LAMPEOLJE"
        ordliste(23) = "HÅNDBOK"
        ordliste(24) = "HÅNDKREM"
        ordliste(25) = "ARBEIDSØKT"
        ordliste(26) = "BENKPRESS"
        ordliste(27) = "TIMEPLAN"
        ordliste(28) = "SKOLEKLASSE"
        ordliste(29) = "DØRBLAD"
        ordliste(30) = "PARKETT"
        ordliste(31) = "SPEILBILDE"
        ordliste(32) = "MINNEPENN"
        ordliste(33) = "PROSESSOR"
        ordliste(34) = "HOVEDKORT"
        ordliste(35) = "INFORMATIKK"
        ordliste(36) = "TEKNOLOGI"
        ordliste(37) = "UTDANNING"
        ordliste(38) = "SKOLEGANG"
        ordliste(39) = "SKRIVEPULT"
        ordliste(40) = "BRØDSKIVE"
        ordliste(41) = "BILMERKE"
        ordliste(42) = "MUSMATTE"
        ordliste(43) = "BOKHYLLE"
        ordliste(44) = "LYSPÆRE"
        ordliste(45) = "STIKKONTAKT"
        ordliste(46) = "RINGPERM"
        ordliste(47) = "LISTVERK"
        ordliste(48) = "PLANKE"
        ordliste(49) = "HAMMER"
        ordliste(50) = "SAGBLAD"
        ordliste(51) = "TAKRENNE"
        ordliste(52) = "SPRINGVANN"
        ordliste(53) = "ØLGLASS"
        ordliste(54) = "MALTØL"
        ordliste(55) = "JULEØL"
        ordliste(56) = "BRYGGERI"
        ordliste(57) = "GJÆRINGSTANK"
        ordliste(58) = "VASKEMASKIN"
        ordliste(59) = "KLESVASK"
        ordliste(60) = "BUNAD"
        ordliste(61) = "SLIPS"
        ordliste(62) = "STIKKPRØVE"
        ordliste(63) = "SYKEHUS"
        ordliste(64) = "AVTALE"
        ordliste(65) = "BILDEKK"
        ordliste(66) = "EGENSKAPER"
        ordliste(67) = "VERDI"
        ordliste(68) = "NØYAKTIG"
        ordliste(69) = "VIKTIGPER"
        ordliste(70) = "AVANSERT"
        ordliste(71) = "KORSTOG"
        ordliste(72) = "ABBEDISSEN"
        ordliste(73) = "ABBEDISSER"
        ordliste(74) = "ABBREVIERE"
        ordliste(75) = "ABBREVIERT"
        ordliste(76) = "ABDIKASJON"
        ordliste(77) = "ABDISERING"
        ordliste(78) = "ABDOMENENE"
        ordliste(79) = "ABDOMINALE"
        ordliste(80) = "ABDOMINALT"
        ordliste(81) = "ABERRASJON"
        ordliste(82) = "ABSTRAHERE"
        ordliste(83) = "ABSTRAHERT"
        ordliste(84) = "ABSTRAKTER"
        ordliste(85) = "ABSTRAKTET"
        ordliste(86) = "ABSURDISME"
        ordliste(87) = "ABSURDITET"
        ordliste(88) = "ACETYLENET"
        ordliste(89) = "ACTIONFILM"
        ordliste(90) = "ACTIONFYLT"
        ordliste(91) = "ADAMSDRAKT"
        ordliste(92) = "ADAMSEPLER"
        ordliste(93) = "ADAMSEPLET"
        ordliste(94) = "ADAMSNATUR"
        ordliste(95) = "ADAMSSLEKT"
        ordliste(96) = "ADAPSJONEN"
        ordliste(97) = "ADAPSJONER"
        ordliste(98) = "ADAPTASJON"
        ordliste(99) = "ADAPTERING"
        ordliste(100) = "ADDENDAENE"
        ordliste(101) = "ADDENDUMET"
        ordliste(102) = "ADDERINGEN"
        ordliste(103) = "ADDERINGER"
        ordliste(104) = "ADDISJONEN"
        ordliste(105) = "ADDISJONER"
        ordliste(106) = "ADDITIVENE"
        ordliste(107) = "ADEKVANSEN"
        ordliste(108) = "ADEKVANSER"
        ordliste(109) = "ADELSBREVA"
        ordliste(110) = "ADELSKAPEN"
        ordliste(111) = "ADELSKAPER"
        ordliste(112) = "BÅTUTSTYR"
        ordliste(113) = "BÅTULYKKE"
        ordliste(114) = "BÅTULYKKA"
        ordliste(115) = "BÅTTYVERI"
        ordliste(116) = "BÅTTYVENE"
        ordliste(117) = "BØYNINGER"
        ordliste(118) = "BØYNINGEN"
        ordliste(119) = "BØYLELÅSA"
        ordliste(120) = "BØYLEHEST"
        ordliste(121) = "BØYKINGEN"
        ordliste(122) = "BØYINGENE"
        ordliste(123) = "CELLOFAN"
        ordliste(124) = "CELLOENE"
        ordliste(125) = "COVERLÅT"
        ordliste(126) = "CHEMISER"
        ordliste(127) = "COWBOYEN"
        ordliste(128) = "DÅRSKAP"
        ordliste(129) = "DÅRLIGE"
        ordliste(130) = "DÅRINGA"
        ordliste(131) = "DÅRENDE"
        ordliste(132) = "DÅREHUS"
        ordliste(133) = "DÅPSLUE"
        ordliste(134) = "DÅPSLUA"
        ordliste(135) = "DÅPSDAG"
        ordliste(136) = "DÅNINGA"
        ordliste(137) = "DÅNENDE"
        ordliste(138) = "DÅHJORT"
        ordliste(139) = "DÅDYRET"
        ordliste(140) = "DÅDRIKT"
        ordliste(141) = "DÅDRIKE"
        ordliste(142) = "DÅDLØST"
        ordliste(143) = "DÅDLØSE"
        ordliste(144) = "DÅDLAUS"
        ordliste(145) = "DØYVSLE"
        ordliste(146) = "DØYVSLA"
        ordliste(147) = "DØDNING"
        ordliste(148) = "DØDMANN"
        ordliste(149) = "DØDGANG"
        ordliste(150) = "DØDFØDT"
        ordliste(151) = "DØDENDE"
        ordliste(152) = "DØDELIG"
        ordliste(153) = "DØDBLEK"
        ordliste(154) = "DØDBALL"
        ordliste(155) = "DØBLING"
        ordliste(156) = "DØBLETE"
        ordliste(157) = "DØBLEDE"
        ordliste(158) = "EVNUKK"
        ordliste(159) = "EVNING"
        ordliste(160) = "EVNETE"
        ordliste(161) = "EVNENE"
        ordliste(162) = "EVNEDE"
        ordliste(163) = "EVLING"
        ordliste(164) = "EVJENE"
        ordliste(165) = "EVENDE"
        ordliste(166) = "EVASIV"
        ordliste(167) = "EVAENE"
        ordliste(168) = "EUTROF"
        ordliste(169) = "EUROER"
        ordliste(170) = "EUROEN"
        ordliste(171) = "EUNUKK"
        ordliste(172) = "EUFORI"
        ordliste(173) = "EUFONI"
        ordliste(174) = "ETYNET"
        ordliste(175) = "ETYLET"
        ordliste(176) = "FÅRET"
        ordliste(177) = "FØRTE"
        ordliste(178) = "FØRST"
        ordliste(179) = "FØRET"
        ordliste(180) = "FØRES"
        ordliste(181) = "FØRER"
        ordliste(182) = "FØNTE"
        ordliste(183) = "FØNET"
        ordliste(184) = "FØNES"
        ordliste(185) = "FØNER"
        ordliste(186) = "FØNEN"
        ordliste(187) = "FØLTE"
        ordliste(188) = "FØLLE"
        ordliste(189) = "FØLLA"
        ordliste(190) = "FØLGE"
        ordliste(191) = "GÅTT"
        ordliste(192) = "GÅTE"
        ordliste(193) = "GÅTA"
        ordliste(194) = "GÅSA"
        ordliste(195) = "GÅRE"
        ordliste(196) = "GÅRD"
        ordliste(197) = "GÅRA"
        ordliste(198) = "GÅET"
        ordliste(199) = "GÅEN"
        ordliste(200) = "GØYV"
        ordliste(201) = "GØYS"
        ordliste(202) = "GØYE"
        ordliste(203) = "GØYA"
        ordliste(204) = "GØRR"
        ordliste(205) = "GØLV"
        ordliste(206) = "GÆRN"
        ordliste(207) = "GYVE"
        ordliste(208) = "GYVD"
        ordliste(209) = "GYVA"
        ordliste(210) = "GYTT"
        ordliste(211) = "GYTE"
        ordliste(212) = "GYTA"
        ordliste(213) = "GYST"
        ordliste(214) = "GYSS"
        ordliste(215) = "GYSJ"
        ordliste(216) = "GYSE"
        ordliste(217) = "GRØN"
        ordliste(218) = "GRÆL"
        ordliste(219) = "GRYS"
        ordliste(220) = "GRYR"
        ordliste(221) = "GRYN"
        ordliste(222) = "GRYL"
        ordliste(223) = "GRYA"
        ordliste(224) = "GRUT"
        ordliste(225) = "GRUS"
        ordliste(226) = "GRUM"
        ordliste(227) = "GRUE"
        ordliste(228) = "GRUA"
        ordliste(229) = "GROV"
        ordliste(230) = "GROS"
        ordliste(231) = "GROR"
        ordliste(232) = "HÅV"
        ordliste(233) = "HÅS"
        ordliste(234) = "HÅR"
        ordliste(235) = "HÅP"
        ordliste(236) = "HÅN"
        ordliste(237) = "HÅL"
        ordliste(238) = "HÅK"
        ordliste(239) = "HÅD"
        ordliste(240) = "HÅA"
        ordliste(241) = "HØY"
        ordliste(242) = "HØV"
        ordliste(243) = "HØR"
        ordliste(244) = "HØN"
        ordliste(245) = "HØL"
        ordliste(246) = "HØH"
        ordliste(247) = "HØG"
        ordliste(248) = "HØA"
        ordliste(249) = "HÆV"
        ordliste(250) = "IS"
        ordliste(251) = "IVRE"
        ordliste(252) = "IVRA"
        ordliste(253) = "IVER"
        ordliste(254) = "ITLE"
        ordliste(255) = "ITEM"
        ordliste(256) = "ISÆR"
        ordliste(257) = "ISTE"
        ordliste(258) = "ISSE"
        ordliste(259) = "ISSA"
        ordliste(260) = "ISPE"
        ordliste(261) = "ISOP"
        ordliste(262) = "ISNE"
        ordliste(263) = "ISNA"
        ordliste(264) = "ISME"
        ordliste(265) = "ISLA"
        ordliste(266) = "ISJE"
        ordliste(267) = "ISJA"
        ordliste(268) = "ISET"
        ordliste(269) = "ISES"
        ordliste(270) = "ISER"
        ordliste(271) = "JÅSS"
        ordliste(272) = "JÅSE"
        ordliste(273) = "JÅSA"
        ordliste(274) = "JÅLT"
        ordliste(275) = "JÅLE"
        ordliste(276) = "JÅLA"
        ordliste(277) = "JØYE"
        ordliste(278) = "JØYA"
        ordliste(279) = "JØSS"
        ordliste(280) = "JØDE"
        ordliste(281) = "JYSK"
        ordliste(282) = "JYDE"
        ordliste(283) = "JUVA"
        ordliste(284) = "JUTE"
        ordliste(285) = "JUST"
        ordliste(286) = "JUSS"
        ordliste(287) = "JURY"
        ordliste(288) = "KÅSØR"
        ordliste(289) = "KØLAPP"
        ordliste(290) = "KÅLÅKER"
        ordliste(291) = "KÅLSAFT"
        ordliste(292) = "KÅLROTA"
        ordliste(293) = "KÅLRABI"
        ordliste(294) = "KÅLMARK"
        ordliste(295) = "KÅLMAKK"
        ordliste(296) = "KÅLHODE"
        ordliste(297) = "KÅLHODA"
        ordliste(298) = "LÅVETAK"
        ordliste(299) = "LÅVEDØR"
        ordliste(300) = "LÅVEBRU"
        ordliste(301) = "LÅVEBRO"
        ordliste(302) = "LÅTTENE"
        ordliste(303) = "LÅTINGA"
        ordliste(304) = "LÅRHØNE"
        ordliste(305) = "LÅRHØNA"
        ordliste(306) = "LÅRHALS"
        ordliste(307) = "LÅRENDE"
        ordliste(308) = "LÅRBENA"
        ordliste(309) = "MÅTELIGE"
        ordliste(310) = "MÅTEHOLD"
        ordliste(311) = "MÅSEUNGE"
        ordliste(312) = "MÅSEFUGL"
        ordliste(313) = "MÅSEFJØR"
        ordliste(314) = "MÅSEFJÆR"
        ordliste(315) = "MÅSEEGGA"
        ordliste(316) = "MÅRSKINN"
        ordliste(317) = "MÅRDYRET"
        ordliste(318) = "MÅPINGER"
        ordliste(319) = "MÅPINGEN"
        ordliste(320) = "MÅNEÅRET"
        ordliste(321) = "MÅLVAKTA"
        ordliste(322) = "MÅLTØRKE"
        ordliste(323) = "MÅLTØRKA"
        ordliste(324) = "NÅVÆRENDE"
        ordliste(325) = "NÅDESTØTA"
        ordliste(326) = "NÅDESKUDD"
        ordliste(327) = "NÅDEGAVER"
        ordliste(328) = "NÅDEGAVEN"
        ordliste(329) = "NØTTEOLJE"
        ordliste(330) = "OZONHULLET"
        ordliste(331) = "OVNSFERSKT"
        ordliste(332) = "OVNSFERSKE"
        ordliste(333) = "OVNSDØRENE"
        ordliste(334) = "OVNSBAKING"
        ordliste(335) = "OVERØSINGA"
        ordliste(336) = "OVERØSENDE"
        ordliste(337) = "OVERVÅKING"
        ordliste(338) = "OVERVÅKETE"
        ordliste(339) = "OVERVÅKERE"
        ordliste(340) = "OVERVÅKEDE"
        ordliste(341) = "OVERVÆRING"
        ordliste(342) = "OVERVÆRENE"
        ordliste(343) = "OVERVURDER"
        ordliste(344) = "OVERVUNNET"
        ordliste(345) = "PÅVIRKINGA"
        ordliste(346) = "PÅVIRKENDE"
        ordliste(347) = "PÅVIRKELIG"
        ordliste(348) = "PÅVENTINGA"
        ordliste(349) = "PÅVENTENDE"
        ordliste(350) = "PÅTVINGING"
        ordliste(351) = "PÅTRYKKING"
        ordliste(352) = "PÅTRYKKETE"
        ordliste(353) = "PÅTRYKKENE"
        ordliste(354) = "PÅTRYKKEDE"
        ordliste(355) = "PÅTREFFING"
        ordliste(356) = "PÅTENNENDE"
        ordliste(357) = "PÅTENKINGA"
        ordliste(358) = "PÅTENKENDE"
        ordliste(359) = "PÅTEGNINGA"
        ordliste(360) = "PÅTEGNENDE"
        ordliste(361) = "QUISLINGEN"
        ordliste(362) = "QUADRICEPS"
        ordliste(363) = "QUIZER"
        ordliste(364) = "QUIZEN"
        ordliste(365) = "QUIPUA"
        ordliste(366) = "QUARTO"
        ordliste(367) = "QUARTS"
        ordliste(368) = "QUIPU"
        ordliste(369) = "QUART"
        ordliste(370) = "QUIZ"
        ordliste(371) = "RØYKSOPPER"
        ordliste(372) = "RØYKSOPPEN"
        ordliste(373) = "RØYKSKYENE"
        ordliste(374) = "RØYKSKADET"
        ordliste(375) = "RØYKSKADER"
        ordliste(376) = "RØYKSKADEN"
        ordliste(377) = "SYTØY"
        ordliste(378) = "SYTTI"
        ordliste(379) = "SYTTE"
        ordliste(380) = "SYTNE"
        ordliste(381) = "SYTET"
        ordliste(382) = "SYTES"
        ordliste(383) = "SYTER"
        ordliste(384) = "SYTEN"
        ordliste(385) = "SYSLE"
        ordliste(386) = "SYSLA"
        ordliste(387) = "TÅREN"
        ordliste(388) = "TÅPNE"
        ordliste(389) = "TÅPER"
        ordliste(390) = "TÅPEN"
        ordliste(391) = "TÅLTE"
        ordliste(392) = "TÅLIG"
        ordliste(393) = "TÅLES"
        ordliste(394) = "TÅLER"
        ordliste(395) = "TÅKET"
        ordliste(396) = "TÅKES"
        ordliste(397) = "TÅKER"
        ordliste(398) = "TÅKEN"
        ordliste(399) = "TÅGET"
        ordliste(400) = "TÅGEN"
        ordliste(401) = "TÅFIS"
        ordliste(402) = "TÅENT"
        ordliste(403) = "TÅENE"
        ordliste(404) = "TØYTE"
        ordliste(405) = "UÅRET"
        ordliste(406) = "UÅPNA"
        ordliste(407) = "UØVET"
        ordliste(408) = "UØVDE"
        ordliste(409) = "UVÆRA"
        ordliste(410) = "UVULA"
        ordliste(411) = "UVISS"
        ordliste(412) = "UVETT"
        ordliste(413) = "UVERK"
        ordliste(414) = "UVENN"
        ordliste(415) = "UVANT"
        ordliste(416) = "UVANE"
        ordliste(417) = "UTÅND"
        ordliste(418) = "UTØYA"
        ordliste(419) = "VÅKER"
        ordliste(420) = "VÅKEN"
        ordliste(421) = "VÅGET"
        ordliste(422) = "WIENERBRØD"
        ordliste(423) = "WEB"
        ordliste(424) = "XENOFOBIEN"
        ordliste(425) = "XYLOFONIST"
        ordliste(426) = "XYLOFONEN"
        ordliste(427) = "YNGSTESØNN"
        ordliste(428) = "YNGLETIDEN"
        ordliste(429) = "YNGSTEDAMA"
        ordliste(430) = "YNGSTEBROR"
        ordliste(431) = "YNGSTEBARN"
        ordliste(432) = "YRKESSKOLE"
        ordliste(433) = "YNGREBØLGA"
        ordliste(434) = "ZIMBABWISK"
        ordliste(435) = "ZOOGRAFENE"
        ordliste(436) = "ZOOLOGISKE"
        ordliste(437) = "ZEPPELINER"
        ordliste(438) = "ZOOLOGENE"
        ordliste(439) = "ZOOLOGIEN"
        ordliste(440) = "ZOOLOGIER"
        ordliste(441) = "ZOOLOGISK"
        ordliste(442) = "ÆRESDOKTOR"
        ordliste(443) = "ÆRESKJELLE"
        ordliste(444) = "ÆRESDIPLOM"
        ordliste(445) = "ÆRESBORGER"
        ordliste(446) = "ÆRESBEGREP"
        ordliste(447) = "ØLTAPPERNE"
        ordliste(448) = "ØKE"
        ordliste(449) = "ØYA"
        ordliste(450) = "ØR"
        ordliste(451) = "ÅK"
        ordliste(452) = "ÅL"
        ordliste(453) = "ÅT"
        ordliste(454) = "ÅSYN"
        ordliste(455) = "ÅNDSFRENDE"
        ordliste(456) = "ÅNDSFRIDOM"
        ordliste(457) = "ÅNDSFRIHET"
        ordliste(458) = "ÅNDSFRISKE"
        ordliste(459) = "ÅNDSARBEID"
        ordliste(460) = "ÅNDSARVENE"
        ordliste(461) = "ÅNDSBROREN"
        ordliste(462) = "AVVIK"
        ordliste(463) = "AVVEK"
        ordliste(464) = "AVVEI"
        ordliste(465) = "AVVEG"
        ordliste(466) = "AVSLÅ"
        ordliste(467) = "AVSLO"
        ordliste(468) = "AVSKY"


    End Sub

    'Prosedyre som skjuler alle labels, knapper o.l.
    Private Sub skjulAlt()
        For i = 1 To 10 'Skjuler alle knapper
            Me.Controls("Button" & i).Visible = False
        Next

        For i = 1 To 40 'Skjuler alle labels
            Me.Controls("Label" & i).Visible = False
        Next

        For i = 1 To 17 'Skjuler alle pictureboxes
            Me.Controls("Picturebox" & i).Visible = False
        Next


    End Sub

    'Prosedyre som viser hovedmeny
    Private Sub visHovedmeny()
        For i = 6 To 9 'Viser alle knapper
            Me.Controls("Button" & i).Visible = True
        Next

        Label40.Visible = True 'viser overskrift
    End Sub


    'Prosedyre som viser et tilfeldig "slapstick"-bilde
    Private Sub slapstickbilde()
        Dim bildenr As Integer
        Randomize()
        bildenr = CInt(Int((8 * Rnd()) + 10))
        PictureBox9.Visible = False
        If spillAvLyd = True Then
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("slapstick" & bildenr - 9 & 1), AudioPlayMode.Background)
        End If
        Me.Controls("PictureBox" & bildenr).Visible = True
    End Sub

    'Prosedyre som nullstiller bildevisning og viser "tomt" bilde
    Private Sub nullstillBilde()
        For i = 1 To 17
            Me.Controls("PictureBox" & i).Visible = False
        Next
        PictureBox1.Visible = True
    End Sub

    'Prosedyre som viser knapper o.l. til enspillerskjermbilde
    Private Sub visEnspillerskjermbilde()
        Button3.Visible = True
        Button5.Visible = True
    End Sub

    'Prosedyre som viser knapper o.l. til tospillerskjermbilde
    Private Sub visTospillerskjermbilde()
        Button3.Visible = True
        Button10.Visible = True
    End Sub

    'Prosedyre som spiller tilfeldig valgt krittlyd
    Private Sub krittlyd()
        If spillAvLyd = True Then
        Dim lydnr As Integer
        Randomize()
        lydnr = CInt(Int((15 * Rnd()) + 1))
        My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("kritt" & lydnr), AudioPlayMode.Background)
        End if
    End Sub

    'Funksjon som kjøres hver gang en bokstav i bokstavskyen klikkes på
    Private Function bokstavsjekk(ByVal valgtBokstav As String, valgtLabel As Integer)
        krittlyd() 'spiller krittlyd

        'Fjerner bokstaven fra bokstavsky
        Me.Controls("label" & valgtLabel).Visible = False
        'definerer variabelen antallForekomster
        Dim antallForekomster As Integer = 0
        'sjekker om noen av bokstaven finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = valgtBokstav Then
                Me.Controls("label" & i).Text = valgtBokstav
                antallForekomster = antallForekomster + 1
            End If
        Next
        'oppdaterer antall bokstaver som gjenstår før spillet er vunnet
        bokstaverIgjen = bokstaverIgjen - antallForekomster

        'sjekker om spilleren har vunnet
        If bokstaverIgjen = 0 Then
            'fjerner bokstavskyen
            For i = 11 To 39 'skal skjule label 11 til og med label 39
                Me.Controls("label" & i).Visible = False
            Next
            MsgBox("Du har vunnet!") 'viser vinnerskjermbilde
        End If

        'sjekker om spilleren skal miste liv
        If antallForekomster = 0 Then
            liv = liv - 1
        End If

        'sjekker om spilleren er død/hengt
        Select Case liv
            Case 8
                PictureBox1.Visible = True
            Case 7
                PictureBox2.Visible = True
                PictureBox1.Visible = False
            Case 6
                PictureBox3.Visible = True
                PictureBox2.Visible = False
            Case 5
                PictureBox4.Visible = True
                PictureBox3.Visible = False
            Case 4
                PictureBox5.Visible = True
                PictureBox4.Visible = False
            Case 3
                PictureBox6.Visible = True
                PictureBox5.Visible = False
            Case 2
                PictureBox7.Visible = True
                PictureBox6.Visible = False
            Case 1
                PictureBox8.Visible = True
                PictureBox7.Visible = False
            Case 0
                PictureBox9.Visible = True
                PictureBox8.Visible = False
                MsgBox("Du har tapt og har blitt hengt..")
                skjulBokstavsky()
                slapstickbilde()
        End Select

    End Function

    'Prosedyre som viser bokstaver i bokstavsky
    Private Sub visBokstavskybokstaver()
        For i = 11 To 39 'skal vise label 11 til og med label 39
            Me.Controls("label" & i).Visible = True
        Next
    End Sub

    'Prosedyre som viser underscores basert på spillordlengde
    Private Sub visUnderscores()
        For i = 1 To 10 'skjuler alle
            Me.Controls("label" & i).Visible = False
        Next

        For i = 1 To spillordlengde 'viser riktig antall
            Me.Controls("label" & i).Visible = True
        Next
    End Sub

    'Prosedyre som skjuler bokstaver i bokstavsky
    Private Sub skjulBokstavsky()
        For i = 11 To 39 'skal vise label 11 til og med label 39
            Me.Controls("label" & i).Visible = False
        Next
    End Sub

    'Prosedyre som bytter ut bokstaver med underscores
    Private Sub fjernBokstaverVisUnderscores()
        For i = 1 To 10
            Me.Controls("label" & i).Text = "_"
        Next
    End Sub

    'Prosedyre som starter nytt spill for en spiller
    Private Sub startEnspiller()
        skjulAlt()
        'lastInnOrdliste()
        visEnspillerskjermbilde()
        nullstillBilde() 'nullstiller bilde
        visBokstavskybokstaver() 'Viser bokstaver i bokstavsky

        'Sjekker lengden på ordlisten og laster verdien inn i variabel
        Dim ordlistelengde As Integer
        ordlistelengde = ordliste.Length
        'DEBUG MsgBox("ordlistelengde: " & ordlistelengde)

        'Velger tilfeldig tall mellom 0 og ordlistelengde og laster verdien inn i variabel
        Randomize()
        Dim tilfeldig As Integer = CInt(Int((ordlistelengde * Rnd()) + 1)) 'xxx Tror dette er ok nå. Her var det noe feil som gjorder at "tilfeldig" kunne bli lik null.
        'DEBUG MsgBox("tilfeldig tall: " & tilfeldig)

        'Henter inn ord basert på det tilfeldige tallet
        Dim spillord As String = ordliste(tilfeldig - 1)

        spillordlengde = spillord.Length
        'DEBUG MsgBox("Spillord: " & spillord)

        spillmatrise = spillord.ToCharArray
        'DEBUG MsgBox("spillmatrise: " & spillmatrise(0))

        'legger inn underscores/"nullstiller" i alle labels
        fjernBokstaverVisUnderscores()

        'viser underscores basert på spillordlengde
        visUnderscores()

        'setter variabelen "liv" til 8 (som tilsvarer maksimalt antall streker på tegningen)
        liv = 8
        'DEBUG MsgBox("liv = " & liv)

        'setter variabelen bokstaverIgjen lik lengden på spillordet
        bokstaverIgjen = spillordlengde
        'DEBUG MsgBox("Antall bokstaver igjen å finne/bokstaverIgjen " & bokstaverIgjen)
    End Sub

    'Prosedyre som starter nytt spill for to spillere
    Private Sub startTospiller()
        'skjuler knapper o.l.
        skjulAlt()

        'viser tospillerskjermbilde
        visTospillerskjermbilde()

        'viser tomt bilde / bakgrunnsbilde
        nullstillBilde()

        'Viser bokstaver i bokstavsky
        visBokstavskybokstaver()

        'Lar motspiller taste inn spillordet i en inputbox
        'konverterer til store bokstaver
        'sjekker at ordet er godkjent (1-10 bokstaver, A-Å, ingen spesialtegn eller tall)
        Dim spillord As String
        Dim godkjent As Boolean
        Dim antallGodkjenteBokstaver
        godkjent = False

        Do
            antallGodkjenteBokstaver = 0

            Do 'løkke: sjekker lengden på ordet
                spillord = InputBox("Skriv inn ord. Maks 10 bokstaver. A-Å").ToUpper
            Loop Until 0 < spillord.Length And spillord.Length < 11

            spillmatrise = spillord.ToCharArray 'laster inn ord i matrise

            For i = 0 To (spillord.Length - 1) 'løkke: kontrollerer at hver bokstav er i området A-Å  XXX fungerer ikke med Æ og Ø..
                godkjent = spillmatrise(i) Like "[A-Å]"
                If spillmatrise(i) = "Æ" Or spillmatrise(i) = "Ø" Then 'Nødløsning for å få godkjent bokstavene Æ og Ø
                    godkjent = True
                End If
                ListBox1.Items.Add("bokstav nr " & (i + 1) & " : " & spillmatrise(i) & " " & godkjent) 'XXX til debug-bruk
                If godkjent = True Then
                    antallGodkjenteBokstaver = antallGodkjenteBokstaver + 1
                End If
            Next

        Loop Until antallGodkjenteBokstaver = spillord.Length



        spillordlengde = spillord.Length

        'DEBUG MsgBox("Spillord: " & spillord)

        'DEBUG MsgBox("spillmatrise: " & spillmatrise(0))

        'legger inn underscores/"nullstiller" i alle labels
        fjernBokstaverVisUnderscores()
        'viser underscores basert på spillordlengde
        visUnderscores()
        'setter variabelen "liv" til 8 (som tilsvarer maksimalt antall streker på tegningen)
        liv = 8
        'DEBUG MsgBox("liv = " & liv)
        'setter variabelen bokstaverIgjen lik lengden på spillordet
        bokstaverIgjen = spillordlengde
        'DEBUG MsgBox("Antall bokstaver igjen å finne/bokstaverIgjen " & bokstaverIgjen)
    End Sub

    'Laster inn ord i matrise og bokstaver i bokstavsky. Viser tomt bilde.
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("Denne knappen gjør ingenting..")
        'Laster inn ord i ordlistematrisen
        'lastInnOrdliste()

        'viser tomt bilde / bakgrunnsbilde
        'nullstillBilde()
    End Sub

    'GAMMEL Starter spill for en spiller
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("Denne knappen gjør ingenting")


    End Sub

    'Starter spill for en spiller (fra hovedmeny)
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        startEnspiller()
    End Sub

    'Starter nytt spill for en spiller (fra enspillerskjermbilde)
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        startEnspiller()
    End Sub

    'Starter spill for to spillere (fra hovedmeny)
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        startTospiller()
    End Sub

    'Starter nytt spill for to spillere (fra tospillerskjermbilde)
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        startTospiller()
    End Sub



    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        bokstavsjekk("A", 11)
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        bokstavsjekk("B", 12)
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        bokstavsjekk("C", 13)
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        bokstavsjekk("D", 14)
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        bokstavsjekk("E", 15)
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        bokstavsjekk("F", 16)
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        bokstavsjekk("G", 17)
    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        bokstavsjekk("H", 18)
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        bokstavsjekk("I", 19)
    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        bokstavsjekk("J", 20)
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        bokstavsjekk("K", 21)
    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click
        bokstavsjekk("L", 22)
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        bokstavsjekk("M", 23)
    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        bokstavsjekk("N", 24)
    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        bokstavsjekk("O", 25)
    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click
        bokstavsjekk("P", 26)
    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        bokstavsjekk("Q", 27)
    End Sub

    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        bokstavsjekk("R", 28)
    End Sub

    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click
        bokstavsjekk("S", 29)
    End Sub

    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        bokstavsjekk("T", 30)
    End Sub

    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click
        bokstavsjekk("U", 31)
    End Sub

    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click
        bokstavsjekk("V", 32)
    End Sub

    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        bokstavsjekk("W", 33)
    End Sub

    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        bokstavsjekk("X", 34)
    End Sub

    Private Sub Label35_Click(sender As Object, e As EventArgs) Handles Label35.Click
        bokstavsjekk("Y", 35)
    End Sub

    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        bokstavsjekk("Z", 36)
    End Sub

    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.Click
        bokstavsjekk("Æ", 37)
    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click
        bokstavsjekk("Ø", 38)
    End Sub

    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click
        bokstavsjekk("Å", 39)
    End Sub


    Private Sub btn2player_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("denne knappen gjør ingenting..")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        skjulAlt()
        visHovedmeny()
    End Sub


    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        End 'Avslutter programmet
    End Sub




    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If spillAvLyd = True Then
            spillAvLyd = False
            Button8.Text = "Slå på lyd"
        ElseIf spillAvLyd = False Then
            spillAvLyd = True
            Button8.Text = "Slå av lyd"
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        lastInnOrdliste()
        spillAvLyd = True
        Button11.Visible = False
        visHovedmeny()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
