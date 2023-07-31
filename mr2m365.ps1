#------------------------------------------------------------------------------------------------------------------
#
# MIT License
#
# Copyright (c) 2023 Sarunas Koncius
#
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
# documentation files (the "Software"), to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and
# to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of
# the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
# THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
# CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.
#
#------------------------------------------------------------------------------------------------------------------
#
# PowerShell Source Code
#
# NAME:
#    mr2m365.ps1
#
# AUTHOR:
#    Sarunas Koncius
#
# VERSION:
# 	 0.7.8
#
# MODIFIED:
#	 2023-07-31
#
#------------------------------------------------------------------------------------------------------------------


<#
	.SYNOPSIS
        PowerShell skriptas sukuria ir atnaujina mokyklos Microsoft 365 aplinkos vartotojų paskyras pagal informaciją,
        esančią Mokinių registre.

	.DESCRIPTION
        ???

	.NOTES
        ???

#>


#
function ReplaceNonLatinCharacters
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}


#
function Rodyti-meniu {
    param (
        [string]$Pavadinimas = 'm r 2 m 3 6 5'
    )
    Clear-Host
    Write-Host "-----------------------------------------------------------------------------------------"
    Write-Host "                                       $Pavadinimas                                      "
    Write-Host "-----------------------------------------------------------------------------------------"
    Write-Host "Prisijungta prie Microsoft 365 aplinkos: " -NoNewline
    Write-Host $MG_M365_aplinka -ForegroundColor Cyan
    Write-Host "Prisijungta naudojant vartotojo paskyrą: " -NoNewline
    Write-Host $MG_M365_vartotojas -ForegroundColor Cyan
    Write-Host "Darbinis aplankas: " -NoNewline
    Write-Host (Get-Location).Path -ForegroundColor Cyan
    Write-Host "Besimokančių asmenų sąrašo failas: " -NoNewline
    Write-Host $Mokiniu_saraso_failas_MR -ForegroundColor Cyan
    Write-Host "Darbinio sąrašo failas: " -NoNewline
    Write-Host $Darbinio_saraso_failas -ForegroundColor Cyan
    Write-Host "Pakoreguoto sąrašo failas: " -NoNewline
    Write-Host $Pakoreguoto_saraso_failas -ForegroundColor Cyan
    Write-Host "Nuskaityta įrašų iš besimokančių asmenų sąrašo: " -NoNewline
    Write-Host $Visi_mokiniai_MR.Count -ForegroundColor Cyan
    Write-Host "Nuskaityta mokinių paskyrų iš Microsoft 365 aplinkos: " -NoNewline
    Write-Host $Visi_mokiniai_M365.Count -ForegroundColor Cyan
    Write-Host "Suformuota įrašų darbiniame sąraše: " -NoNewline
    Write-Host $Visi_mokiniai.Count -ForegroundColor Cyan
    Write-Host "Nuskaityta įrašų iš pakoreguoto sąrašo: " -NoNewline
    Write-Host $Pakoreguoti_mokiniai_CSV.Count -ForegroundColor Cyan
    Write-Host "-----------------------------------------------------------------------------------------"
    
    Write-Host "`n1: Prisijungti prie Microsoft Graph API                                     " -NoNewline
    if ($Busena_MG.Length -eq 11) { Write-Host "  " -NoNewline }
    Write-Host $Busena_MG -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`n2: Nuskaityti iš Mokinių registro atsisiųstą besimokančių asmenų sąrašą     " -NoNewline
    if ($Busena_MR.Length -eq 10) { Write-Host "  " -NoNewline }
    Write-Host $Busena_MR -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`n3: Nuskaityti mokinių paskyrų informaciją iš Microsoft 365 aplinkos         " -NoNewline
    if ($Busena_M365.Length -eq 10) { Write-Host "  " -NoNewline }
    Write-Host $Busena_M365 -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`n4: Suformuoti darbinį sąrašą iš visos turimos informacijos                  " -NoNewline
    if ($Busena_DS.Length -eq 10) { Write-Host "  " -NoNewline }
    Write-Host $Busena_DS -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`n5: Patikrinimui išsaugoti darbinio sąrašo duomenis CSV faile                " -NoNewline
    if ($Busena_CSV.Length -eq 9) { Write-Host "  " -NoNewline }
    Write-Host $Busena_CSV -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`n6: Užkrauti pakoreguotą sąrašą iš CSV failo                                 " -NoNewline
    if ($Busena_CSV_mokiniai.Length -eq 10) { Write-Host "  " -NoNewline }
    Write-Host $Busena_CSV_mokiniai -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`n7: Atnaujinti mokinių paskyras Microsoft 365 aplinkoje                      " -NoNewline
    Write-Host $Busena_mokiniai -ForegroundColor DarkBlue -BackgroundColor White
#    Write-Host "`n8: Nuskaityti klasių saugos grupių informacija iš Microsoft 365 aplinkos    " -NoNewline
#    Write-Host $Busena_CSV_klases -ForegroundColor DarkBlue -BackgroundColor White
#    Write-Host "`n9: Atnaujinti klasių saugos grupių paskyras Microsoft 365 aplinkoje         " -NoNewline
#    Write-Host $Busena_klases -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host "`nQ: Baigti darbą"
    Write-Host "`n-----------------------------------------------------------------------------------------"
}


#
# Mokyklos naudojamas interneto domeno vardas
$Domeno_vardas = "eportfelis.net"

#
$MokyklosPavadinimas = "Elektroninio portfelio bandymų mokykla"

#
$MokyklosMiestas = "Kaunas"

# Visuotinio administratoriaus teises turinčios paskyros e. pašto adresas
$VisuotinioAdministratoriausSmtpAdresas = "o365.administratorius@eportfelis.net"

# Mokytojų saugos grupės e. pašto adresas
$GrupesVisiMokytojaiSmtpAdresas = "visi.mokytojai@eportfelis.net"

# Ankstesnieji mokslo metai
$Ankstesnieji_mokslo_metai = "2022-2023"

# Naujieji mokslo metai
$Naujieji_mokslo_metai = "2023-2024"

# 
$Mokiniu_saraso_failas_MR = "besimokantys.csv"

# 
$Darbinio_saraso_failas = "darbinis.csv"

# 
$Pakoreguoto_saraso_failas = "pakoreguotas.csv"

#
$Sukurtu_paskyru_failas = "sukurtos_paskyros.csv"


#
$Busena_MG = "Neprisijungta"
$Busena_MR = "Nenuskaityta"
$Busena_M365 = "Nenuskaityta"
$Busena_DS = "Nesuformuota"
$Busena_CSV = "Neišsaugota"
$Busena_CSV_mokiniai = "Nenuskaityta"
$Busena_mokiniai = "Neatnaujinta"
$Busena_CSV_klases = "Nenuskaityta"
$Busena_klases = "Neatnaujinta"

$MG_M365_aplinka = "?"
$MG_M365_vartotojas = "?"


#
do {
    Rodyti-meniu
    $Pasirinkimas = Read-Host "Pasirinkite veiksmą"

    switch ($Pasirinkimas) {

        '1' {
            # 1: Prisijungti prie Microsoft Graph API
            Clear-Host
            Write-Host "Tikrinamas Microsoft Graph modulis..."
 
            #
            if (Get-InstalledModule Microsoft.Graph) {
            # Connect to MS Graph
            Write-Host "Microsoft Graph modulis yra įdiegtas"
            } else {
            Write-Host "Microsoft Graph modulis nerastas - įdiekite jį" -ForegroundColor Black -BackgroundColor Yellow
            # Install-Module Microsoft.Graph -Scope AllUsers -Force
            exit
            }

            #
            Write-Host "Prisijungiama prie Microsoft Graph API..."
            Connect-Graph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All","Group.ReadWrite.All"
            $MG_informacija = Get-MgContext
            $MG_M365_vartotojas = $MG_informacija.Account
            $MG_M365_aplinka = (Get-MgOrganization).DisplayName
            if ($MG_M365_vartotojas -ne $null) { $Busena_MG = "Prisijungta" } else { $Busena_MG = "Neprisijungta" }

        } '2' {
            # 2: Nuskaityti iš Mokinių registro atsisiųstą besimokančių asmenų sąrašą
            Clear-Host
            Write-Host "Nuskaitomas besimokančių asmenų sąrašas..."

            #
            $Visi_mokiniai_MR = Import-Csv $Mokiniu_saraso_failas_MR -Encoding UTF8 -Delimiter ";"
            Write-Host "Nuskaityta įrašų iš besimokančių asmenų sąrašo:", $Visi_mokiniai_MR.Count
            if ($Visi_mokiniai_MR.Count -ne 0) { $Busena_MR = "Nuskaityta" } else { $Busena_MR = "Nenuskaityta" }

        } '3' {
            # 3: Nuskaityti mokinių paskyrų informaciją iš Microsoft 365 aplinkos
            Clear-Host
            Write-Host "Nuskaitomos mokinių paskyros iš Microsoft 365 aplinkos..."
        
            #
            $Vartotojo_paskyros_laukai_M365 = @(
                'AccountEnabled',
                'AssignedLicenses',
                'AssignedPlans',
                'City',
                'CompanyName',
                'Country',
                'Department',
                'DisplayName',
                'GivenName',
                'Id',
                'JobTitle',
                'EmployeeId',
                'EmployeeType',
                'OfficeLocation',
                'Surname',
                'UserPrincipalName'
            )
            $GetMgUserKlaidos = 0
            $Visi_mokiniai_M365 = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq 314c4481-f395-4525-be8b-2ec4bb1e9d91)" -All -Property $Vartotojo_paskyros_laukai_M365 -ExpandProperty Manager -OrderBy Surname -ErrorVariable GetMgUserKlaidos
            Write-Host "Nuskaityta mokinių paskyrų iš Microsoft 365 aplinkos:", $Visi_mokiniai_M365.Count
            if ($GetMgUserKlaidos.Count -eq 0) { $Busena_M365 = "Nuskaityta" } else { $Busena_M365 = "Nenuskaityta" }

            # https://lazyadmin.nl/powershell/get-mguser/
            # -ExpandProperty Manager, EmployeeOrgData
            # EmployeeOrgData                       : Microsoft.Graph.PowerShell.Models.MicrosoftGraphEmployeeOrgData
            # Get-MgUser -UserId edita.rabizaite@kgm.lt -ExpandProperty manager | Select @{Name = 'Manager'; Expression = {$_.Manager.AdditionalProperties.displayName}}
        
        } '4' {
            Clear-Host
            Write-Host "Formuojamas darbinis sąrašas..."

            # 
            Write-Host "Į darbinį sąrašą perkeliama Microsoft 365 mokinių paskyrų informacija..."
            $Visi_mokiniai = @()
            foreach ($Mokinio_paskyra_M365 in $Visi_mokiniai_M365) {
                $Mokinio_informacija = New-Object PSObject
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Prisijungimo vardas" -Value $Mokinio_paskyra_M365.UserPrincipalName
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Paskyra aktyvi" -Value $Mokinio_paskyra_M365.AccountEnabled
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "ID" -Value $Mokinio_paskyra_M365.Id
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Rodomas vardas" -Value $Mokinio_paskyra_M365.DisplayName
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Vardas" -Value $Mokinio_paskyra_M365.GivenName
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Pavardė" -Value $Mokinio_paskyra_M365.Surname
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Pareigos" -Value $Mokinio_paskyra_M365.JobTitle
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Klasė" -Value $Mokinio_paskyra_M365.Department
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Organizacija" -Value $Mokinio_paskyra_M365.CompanyName
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Biuras" -Value $Mokinio_paskyra_M365.OfficeLocation
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Miestas" -Value $Mokinio_paskyra_M365.City
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Šalis" -Value $Mokinio_paskyra_M365.Country
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Darbuotojo ID" -Value $Mokinio_paskyra_M365.EmployeeId
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Darbuotojo tipas" -Value $Mokinio_paskyra_M365.EmployeeType
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Šaltinis" -Value "M365"
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Veiksmai" -Value ""
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Pavardė_vardas" -Value ""
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Gimimo_data" -Value ""
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Lytis" -Value ""
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Bylos Nr." -Value ""
	            Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Klasė" -Value ""
                $Visi_mokiniai += $Mokinio_informacija
            }
            Write-Host "Į darbinį sąrašą perkelta Microsoft 365 mokinių paskyrų informacija."

            #
            Write-Host "Į darbinį sąrašą perkeliama Moksleivių registro informacija paskyrų atnaujinimui..."
            foreach ($Mokinio_informacija in $Visi_mokiniai) {
                if ($Mokinio_informacija.Vardas.Length -gt 0 -and $Mokinio_informacija.Pavardė.Length -gt 0) {
                    $Surastas_mokinys_MR = @()
                    $Surastas_mokinys_MR = $Visi_mokiniai_MR | Where-Object { $_."Pavardė Vardas" -match $Mokinio_informacija.Vardas -and $_."Pavardė Vardas" -match $Mokinio_informacija.Pavardė }
                    $PaieskosRezultatai = $Surastas_mokinys_MR | Measure-Object
                    if ($PaieskosRezultatai.Count -eq 1) {
                        $Mokinio_informacija.MR_Pavardė_vardas = $Surastas_mokinys_MR.'Pavardė Vardas'
                        $Mokinio_informacija.MR_Gimimo_data = $Surastas_mokinys_MR.'Gimimo data'
                        $Mokinio_informacija.MR_Lytis = $Surastas_mokinys_MR.Lytis
                        $Mokinio_informacija.'MR_Bylos Nr.' = $Surastas_mokinys_MR.'Bylos Nr./Asm. vardinis Nr.'
                        $Mokinio_informacija.MR_Klasė = $Surastas_mokinys_MR.Klasė
                        $Mokinio_informacija.Veiksmai = "Atnaujinti"
                        $Mokinio_informacija.Šaltinis += "+MR"
                    } elseif ($PaieskosRezultatai.Count -gt 1) {
                        $Mokinio_informacija.MR_Pavardė_vardas = "! Patikrinti rankiniu būdu !"
                        $Mokinio_informacija.Veiksmai = "Patikrinti"
                        $Mokinio_informacija.Šaltinis += "+MR"
                    } else {
                        $Mokinio_informacija.Veiksmai = "Deaktyvuoti"
                    }
                }
            }
            Write-Host "Į darbinį sąrašą perkelta Moksleivių registro informacija paskyrų atnaujinimui."

            #
            Write-Host "Į darbinį sąrašą perkeliama Moksleivių registro informacija naujoms paskyroms..."
            foreach ($Mokinio_informacija_MR in $Visi_mokiniai_MR) {
                $Surastas_mokinys_M365 = @()
                $Surastas_mokinys_M365 = $Visi_mokiniai | Where-Object { $_.'MR_Bylos Nr.' -eq $Mokinio_informacija_MR.'Bylos Nr./Asm. vardinis Nr.' }
                $PaieskosRezultatai = $Surastas_mokinys_M365 | Measure-Object
                if ($PaieskosRezultatai.Count -eq 0) {
                    $Mokinio_informacija = New-Object PSObject
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Prisijungimo vardas" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Paskyra aktyvi" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "ID" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Rodomas vardas" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Vardas" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Pavardė" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Pareigos" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Klasė" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Organizacija" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Biuras" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Miestas" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Šalis" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Darbuotojo ID" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Darbuotojo tipas" -Value ""
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Šaltinis" -Value "MR"
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "Veiksmai" -Value "Sukurti"
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Pavardė_vardas" -Value $Mokinio_informacija_MR.'Pavardė Vardas'
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Gimimo_data" -Value $Mokinio_informacija_MR.'Gimimo data'
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Lytis" -Value $Mokinio_informacija_MR.Lytis
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Bylos Nr." -Value $Mokinio_informacija_MR.'Bylos Nr./Asm. vardinis Nr.'
	                Add-Member -InputObject $Mokinio_informacija -MemberType NoteProperty -Name "MR_Klasė" -Value $Mokinio_informacija_MR.Klasė
                    $Visi_mokiniai += $Mokinio_informacija
                }
            }
            Write-Host "Į darbinį sąrašą perkelta Moksleivių registro informacija naujoms paskyroms."

            Write-Host "Darbinio sąrašo formavimas baigtas."
            if ($Visi_mokiniai.Count -ne 0) { $Busena_DS = "Suformuota" } else { $Busena_DS = "Nesuformuota" }
            


        } '5' {
            Clear-Host
            Write-Host "Išsaugomi patikrinimui darbinio sąrašo duomenys CSV faile..."

            #
            $Visi_mokiniai | Export-CSV $Darbinio_saraso_failas -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Host "CSV faile ", $Darbinio_saraso_failas, " išsaugota įrašų: " $Visi_mokiniai.Count
            Write-Host "Išsaugoti patikrinimui darbinio sąrašo duomenys CSV faile" $Darbinio_saraso_failas"."
            if ((Get-ChildItem $Darbinio_saraso_failas).Length -ne 0) { $Busena_CSV = "Išsaugota" } else { $Busena_CSV = "Neišsaugota" }
            

        } '6' {
            Clear-Host
            Write-Host "Nuskaitomas pakoreguotas mokinių sąrašo CSV failas..."

            #
            $Pakoreguoti_mokiniai_CSV = Import-Csv $Pakoreguoto_saraso_failas -Encoding UTF8 -Delimiter ";"
            $Kuriamos_paskyros_CSV = $Pakoreguoti_mokiniai_CSV | Where-Object { $_.Veiksmai -eq "Sukurti" }
            $Atnaujinamos_paskyros_CSV = $Pakoreguoti_mokiniai_CSV | Where-Object { $_.Veiksmai -eq "Atnaujinti" }
            $Deaktyvuojamos_paskyros_CSV = $Pakoreguoti_mokiniai_CSV | Where-Object { $_.Veiksmai -eq "Deaktyvuoti" }

            Write-Host "Nuskaityta įrašų iš pakoreguoto mokinių sąrašo:", $Pakoreguoti_mokiniai_CSV.Count
            Write-Host "Reikia sukurti paskyrų:", $Kuriamos_paskyros_CSV.Count
            Write-Host "Reikia atnaujinti paskyrų:", $Atnaujinamos_paskyros_CSV.Count
            Write-Host "Reikia deaktyvuoti paskyrų:", $Deaktyvuojamos_paskyros_CSV.Count

            Write-Host "Nuskaitytas pakoreguotas mokinių sąrašo CSV failas."
            if ($Pakoreguoti_mokiniai_CSV.Count -ne 0) { $Busena_CSV_mokiniai = "Nuskaityta" } else { $Busena_CSV_mokiniai = "Nenuskaityta" }


        } '7' {
            Clear-Host
            Write-Host "Pradedmas mokinių paskyrų tvarkymas Microsoft 365 aplinkoje..."

            #
            Write-Host "Kuriamos naujos mokinių paskyros..."
            $Licencijos = @(
                @{SkuId = '314c4481-f395-4525-be8b-2ec4bb1e9d91'}
            )
            $Sukurtos_paskyros = @()
            foreach ($Kuriama_paskyra in $Kuriamos_paskyros_CSV) {
                If ($Kuriama_paskyra.MR_Pavardė_vardas.Contains(" ")) { $Pavarde = $Kuriama_paskyra.MR_Pavardė_vardas.Substring(0, $Kuriama_paskyra.MR_Pavardė_vardas.IndexOf(" ")) } else { $Pavarde = $Kuriama_paskyra.MR_Pavardė_vardas }
                $Pavarde = (Get-Culture).textinfo.totitlecase($Pavarde.ToLower())
                If ($Kuriama_paskyra.MR_Pavardė_vardas.Contains(" ")) { $Vardas = $Kuriama_paskyra.MR_Pavardė_vardas.Substring($Kuriama_paskyra.MR_Pavardė_vardas.LastIndexOf(" ")+1,$Kuriama_paskyra.MR_Pavardė_vardas.Length-$Kuriama_paskyra.MR_Pavardė_vardas.LastIndexOf(" ")-1) } else { $Vardas = $Kuriama_paskyra.MR_Pavardė_vardas }
                $Vardas = (Get-Culture).textinfo.totitlecase($Vardas.ToLower())
                $RodomasVardas = $Vardas + " " + $Pavarde
                $Klase = $Kuriama_paskyra.MR_Klasė + " klasė"
                $Pareigos = $Klase + "s mokin"
                if ($Kuriama_paskyra.MR_Lytis -eq "moteris") { 
                    $Pareigos += "ė"
                } else {
                    $Pareigos += "ys"
                }
                $TrumpasisID = (ReplaceNonLatinCharacters $Vardas.ToLower()) + "." + (ReplaceNonLatinCharacters $Pavarde.ToLower())
                $VartotojoID = $TrumpasisID + "@" + $Domeno_vardas
                $Slaptazodis = @(
                    [char]('ABCDEFGHKLMNPRSTUVZ'.ToCharArray() | get-random),
                    [char]('abcdefghkmnprstuvyz'.ToCharArray() | get-random),
                    [char]('abcdefghkmnprstuvyz'.ToCharArray() | get-random),
                    (0..9 | get-random),
                    (0..9 | get-random),
                    (0..9 | get-random),
                    (0..9 | get-random),
                    (0..9 | get-random)
                ) -join ''
                $SlaptazodzioProfilis = @{
                    ForceChangePasswordNextSignIn = $false
                    Password = $Slaptazodis
                }
                New-MgUser -UserPrincipalName $VartotojoID -PasswordProfile $SlaptazodzioProfilis -AccountEnabled -UsageLocation "LT" -MailNickName $TrumpasisID -DisplayName $RodomasVardas -GivenName $Vardas -Surname $Pavarde -JobTitle $Pareigos -Department $Klase -EmployeeId $Kuriama_paskyra.'MR_Bylos Nr.' -OfficeLocation $Naujieji_mokslo_metai -CompanyName $MokyklosPavadinimas -City $MokyklosMiestas -Country "Lithuania" | Out-Null
                Set-MgUserLicense -UserId $VartotojoID -AddLicenses $Licencijos -RemoveLicenses @()
                $Sukurta_paskyra = New-Object PSObject
	            Add-Member -InputObject $Sukurta_paskyra -MemberType NoteProperty -Name "Klasė" -Value $Klase
	            Add-Member -InputObject $Sukurta_paskyra -MemberType NoteProperty -Name "Pavardė_vardas" -Value $Kuriama_paskyra.MR_Pavardė_vardas
	            Add-Member -InputObject $Sukurta_paskyra -MemberType NoteProperty -Name "Prisijungimo_vardas" -Value $VartotojoID
	            Add-Member -InputObject $Sukurta_paskyra -MemberType NoteProperty -Name "Slaptažodis" -Value $Slaptazodis
                $Sukurtos_paskyros += $Sukurta_paskyra
            }
            $Sukurtos_paskyros | Sort-Object -Property Klasė, Pavardė_vardas | Export-CSV $Sukurtu_paskyru_failas -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Host "Naujos mokinių paskyros sukurtos."
            Write-Host "Sukurta naujų mokinių paskyrų:", $Sukurtos_paskyros.Count
            Write-Host "Sukurtų paskyrų informacija išsaugota CSV faile ", $Sukurtu_paskyru_failas

            #
            Write-Host "Atnaujinama mokinių paskyrų informacija..."
            foreach ($Atnaujinama_paskyra in $Atnaujinamos_paskyros_CSV) {
                $NaujaKlase = $Atnaujinama_paskyra.MR_Klasė + " klasė"
                $NaujosPareigos = $NaujaKlase + "s mokin"
                if ($Atnaujinama_paskyra.MR_Lytis -eq "moteris") { 
                    $NaujosPareigos += "ė"
                } else {
                    $NaujosPareigos += "ys"
                }
                Update-MgUser -UserId $Atnaujinama_paskyra.ID -Department $NaujaKlase -JobTitle $NaujosPareigos -EmployeeId $Atnaujinama_paskyra.'MR_Bylos Nr.' -OfficeLocation $Naujieji_mokslo_metai -CompanyName $MokyklosPavadinimas -City $MokyklosMiestas -Country "Lithuania"
            }
            Write-Host "Atnaujinta mokinių paskyrų informacija."
            Write-Host "Atnaujinta paskyrų:", $Atnaujinamos_paskyros_CSV.Count

            #
            Write-Host "Atnaujinamos ir deaktyvuojamos alumnų paskyros..."
            foreach ($Deaktyvuojama_paskyra in $Deaktyvuojamos_paskyros_CSV) {
                Update-MgUser -UserId $Deaktyvuojama_paskyra.ID -JobTitle "Alumnas"
                $UriAdresas = "https://graph.microsoft.com/v1.0/Users/{"+$Deaktyvuojama_paskyra.ID+"}"
                Invoke-MgGraphRequest -Method PATCH -Uri $UriAdresas -Body @{Department = $null}
                Update-MgUser -UserId $Deaktyvuojama_paskyra.ID -AccountEnabled:$false
            }
            Write-Host "Atnaujintos ir deaktyvuotos alumnų paskyros."
            Write-Host "Deaktyvuoti paskyrų:", $Deaktyvuojamos_paskyros_CSV.Count

            Write-Host "Baigtas mokinių paskyrų tvarkymas Microsoft 365 aplinkoje."
            if ($Pakoreguoti_mokiniai_CSV.Count -ne 0) { $Busena_mokiniai = "Atnaujinta" } else { $Busena_mokiniai = "Neatnaujinta" }


#        } '8' {
#            Clear-Host
#            Write-Host "Nuskaitoma klasių saugos grupių informacija iš Microsoft 365 aplinkos..."
#
#        
#        } '9' {
#            Clear-Host
#            Write-Host "Atnaujinama klasių saugos grupių informacija Microsoft 365 aplinkoje..."
#
        } 


    }
    Write-Host
    pause
 }
 until ($Pasirinkimas -eq 'q')
 Disconnect-MgGraph

