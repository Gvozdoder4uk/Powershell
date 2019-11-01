﻿# TASK FOR HHT EMULATOR



Write-Host "=====================
Выберите контур:   ||
|| 1: VRQ          ||
|| 2: VRX          ||
|| 3: VRA          ||
=====================" -BackgroundColor DarkGreen -ForegroundColor Black
$Contur = Read-Host "Выберите контур"

if ($Contur -eq 1)
{
$shopName = Read-Host "Введите магазин"
$S= "S$shopName"

}
elseif ($Contur -eq 2)
{
$shopName = Read-Host "Введите магазин"
$S= "S$shopName"  
}


test-Path C:\NTSwincash\wm\update\

if(test-Path C:\HHT_Emulator -eq "False")
{
    
}

#wm_hwdeviceid.txt
#wm_treeid.txt

$HASH1 = @{
S002="191";`
S011="8";`
S014="30";`
S015="16";`
S016="103";`
S018="21";`
S019="22";`
S020="99";`
S021="9";`
S022="7";`
S023="422";`
S024="101";`
S025="14";`
S026="107";`
S027="17";`
S028="135";`
S029="138";`
S030="140";`
S032="141";`
S034="142";`
S035="143";`
S036="144";`
S038="145";`
S039="20";`
S041="146";`
S042="423";`
S044="147";`
S045="25";`
S046="148";`
S047="406";`
S048="149";`
S050="150";`
S051="151";`
S052="152";`
S053="154";`
S054="153";`
S055="155";`
S056="156";`
S057="157";`
S058="158";`
S059="189";`
S060="159";`
S061="331";`
S062="160";`
S063="161";`
S064="162";`
S065="306";`
S066="163";`
S067="164";`
S068="165";`
S069="166";`
S070="167";`
S071="168";`
S072="594";`
S073="169";`
S074="354";`
S076="170";`
S077="171";`
S078="172";`
S079="173";`
S080="174";`
S081="175";`
S082="176";`
S083="177";`
S084="178";`
S085="602";`
S086="846";`
S087="179";`
S088="773";`
S089="180";`
S090="181";`
S091="598";`
S092="182";`
S093="603";`
S094="409";`
S095="183";`
S096="220";`
S097="595";`
S098="184";`
S100="18";`
S101="314";`
S102="268";`
S103="254";`
S104="433";`
S105="199";`
S106="255";`
S107="628";`
S108="793";`
S110="642";`
S111="200";`
S112="270";`
S113="649";`
S114="351";`
S115="492";`
S116="491";`
S117="490";`
S118="201";`
S119="202";`
S120="326";`
S121="23";`
S122="266";`
S123="197";`
S124="256";`
S125="221";`
S126="404";`
S127="617";`
S128="618";`
S129="489";`
S130="450";`
S131="639";`
S132="656";`
S133="636";`
S134="204";`
S137="224";`
S138="612";`
S139="601";`
S140="456";`
S141="620";`
S142="418";`
S144="724";`
S145="567";`
S146="267";`
S147="637";`
S149="455";`
S150="327";`
S151="341";`
S152="307";`
S154="488";`
S156="336";`
S157="269";`
S158="621";`
S159="249";`
S160="487";`
S161="661";`
S162="460";`
S163="257";`
S164="402";`
S166="726";`
S167="729";`
S168="472";`
S169="672";`
S171="486";`
S173="312";`
S174="548";`
S175="205";`
S176="673";`
S177="485";`
S179="718";`
S180="209";`
S181="736";`
S182="474";`
S183="706";`
S184="241";`
S185="337";`
S186="712";`
S187="259";`
S188="626";`
S189="623";`
S190="206";`
S191="207";`
S192="599";`
S193="650";`
S194="483";`
S195="252";`
S196="607";`
S197="482";`
S198="675";`
S199="738";`
S202="315";`
S203="475";`
S204="609";`
S206="684";`
S207="258";`
S208="228";`
S209="225";`
S212="666";`
S216="794";`
S217="345";`
S218="339";`
S219="643";`
S220="244";`
S221="800";`
S223="686";`
S224="652";`
S225="321";`
S226="597";`
S227="400";`
S228="737";`
S230="208";`
S232="481";`
S233="682";`
S234="691";`
S235="210";`
S236="747";`
S237="318";`
S239="678";`
S240="766";`
S242="262";`
S243="352";`
S244="484";`
S245="754";`
S246="692";`
S247="263";`
S248="333";`
S249="671";`
S251="720";`
S252="222";`
S253="646";`
S254="362";`
S255="702";`
S256="342";`
S257="308";`
S258="309";`
S259="319";`
S260="247";`
S261="493";`
S263="324";`
S264="801";`
S265="358";`
S265_X="429";`
S266="349";`
S267="343";`
S268="755";`
S269="219";`
S271="320";`
S272="749";`
S274="693";`
S275="742";`
S276="303";`
S277="638";`
S278="779";`
S279="725";`
S280="605";`
S281="452";`
S282="619";`
S283="657";`
S284="250";`
S285="611";`
S286="322";`
S287="405";`
S289="732";`
S292="705";`
S293="710";`
S294="721";`
S297="627";`
S298="688";`
S299="348";`
S301="752";`
S302="211";`
S305="741";`
S306="229";`
S308="260";`
S309="698";`
S311="667";`
S313="633";`
S314="432";`
S316="328";`
S319="630";`
S320="764";`
S321="735";`
S322="728";`
S325="703";`
S327="606";`
S328="701";`
S330="264";`
S331="713";`
S332="802";`
S333="624";`
S335="745";`
S337="717";`
S338="655";`
S339="753";`
S343="761";`
S345="716";`
S346="663";`
S347="608";`
S348="408";`
S350="304";`
S351="704";`
S353="647";`
S354="480";`
S355="677";`
S356="644";`
S358="325";`
S359="640";`
S360="645";`
S361="635";`
S363="653";`
S365="569";`
S366="470";`
S367="670";`
S370="373";`
S371="768";`
S372="743";`
S373="346";`
S374="323";`
S377="226";`
S378="676";`
S379="783";`
S380="631";`
S381="613";`
S383="317";`
S384="329";`
S385="245";`
S386="316";`
S387="659";`
S388="795";`
S390="334";`
S391="714";`
S392="740";`
S393="356";`
S395="709";`
S396="759";`
S398="665";`
S399="625";`
S400="654";`
S401="212";`
S402="685";`
S405="251";`
S409="648";`
S410="632";`
S413="629";`
S415="851";`
S416="748";`
S417="243";`
S418="622";`
S419="641";`
S420="577";`
S421="826";`
S422="478";`
S423="674";`
S424="660";`
S425="683";`
S426="681";`
S428="616";`
S429="248";`
S430="411";`
S430_X="430";`
S432="305";`
S433="758";`
S434="310";`
S435="804";`
S436="311";`
S437="457";`
S438="347";`
S439="338";`
S441="479";`
S443="313";`
S444="213";`
S445="750";`
S446="751";`
S447="707";`
S448="699";`
S449="668";`
S450="669";`
S452="330";`
S453="340";`
S454="719";`
S455="730";`
S456="253";`
S457="335";`
S458="765";`
S459="687";`
S461="600";`
S462="604";`
S463="246";`
S464="242";`
S465="214";`
S466="763";`
S467="727";`
S468="410";`
S469="230";`
S470="694";`
S472="261";`
S473="739";`
S475="731";`
S476="565";`
S477="680";`
S478="203";`
S478_="428";`
S479="700";`
S480="658";`
S481="746";`
S482="350";`
S483="715";`
S484="344";`
S485="614";`
S486="854";`
S487="568";`
S488="634";`
S489="756";`
S491="772";`
S492="734";`
S493="679";`
S494="265";`
S495="662";`
S496="615";`
S497="477";`
S498="664";`
}

$GID = $Hash1[$S]

Write-host "Получение GID магазина и запись в файл wm_treeid.txt ="$GID
$GID | Out-File C:\HHT_Emulator\treeid.txt

Write-Host "Введите идентификатор в формате КОМПАНИЯ_ФАМИЛИЯ
Пример: lux_andreev
        kelly_fokin" -BackgroundColor DarkBlue
$ID = Read-Host "Ожидание ввода:" | Out-File C:\HHT_Emulator\wm_hwdeviceid.txt