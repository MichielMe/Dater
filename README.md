# Dater

# Python Script Uitleg

Dit script is geschreven in Python en maakt gebruik van de pandas bibliotheek om data manipulatie en analyse mogelijk te maken. Het script werkt met Excel-bestanden en de functie `query_result` uit het module `querySim`.

## Werking van het Script

1. De functie `query_result` wordt aangeroepen om data op te halen, die wordt opgeslagen in een pandas DataFrame. Deze DataFrame wordt vervolgens opgeslagen in een Excel-bestand genaamd 'srcexcel.xlsx'.

2. Het script leest vervolgens een ander Excel-bestand genaamd 'vullingen.xlsx', waar het de 'MATERIAL ID's uithaalt.

3. De originele DataFrame wordt gefilterd op basis van deze 'MATERIAL ID's. Alleen de rijen met 'MATERIAL ID's die in 'vullingen.xlsx' aanwezig zijn, worden behouden.

4. Het script vraagt de gebruiker om een suffix en een materiaaltype in te voeren. De suffix wordt aan het einde van de titel toegevoegd en het materiaaltype wordt ingesteld als het nieuwe 'MATERIAL TYPE'. In de webapplicatie worden deze invoervelden aangeboden als een dropdown-menu met vaste keuzes.

5. De gewijzigde DataFrame wordt omgezet naar een lijst van woordenboeken.

6. Tot slot wordt de gewijzigde DataFrame opgeslagen in een nieuw Excel-bestand genaamd 'destexcel.xlsx'.

## Benodigde Bestanden

Dit script vereist de volgende bestanden:

- 'vullingen.xlsx': Dit bestand moet in dezelfde directory staan als het script en bevat de 'MATERIAL ID's waarmee de originele DataFrame wordt gefilterd.

## Output

Dit script genereert de volgende output:

- 'srcexcel.xlsx': Een Excel-bestand dat de originele, ongefilterde DataFrame bevat.
- 'destexcel.xlsx': Een Excel-bestand dat de gefilterde en gewijzigde DataFrame bevat.

## Gebruik

Om dit script te gebruiken, voert u het uit in een Python-omgeving. Wanneer het script wordt uitgevoerd, wordt u gevraagd om een suffix en een materiaaltype in te voeren. Deze waarden worden gebruikt om de 'TITLE' en 'MATERIAL TYPE' kolommen in de DataFrame te wijzigen.

## Afhankelijkheden

Dit script maakt gebruik van de volgende Python-bibliotheken:

- pandas
- openpyxl (nodig voor het lezen en schrijven van Excel-bestanden met pandas)
