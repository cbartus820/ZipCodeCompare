# ZipCodeCompare
The VBA script CompareShipToState.bas is designed to compare ship-to state data from a transaction detail report of sales transactions (in the Sage accounting system) with validated address data based on zip codes transmitted to Vertex. The goal is to identify discrepancies in the Sage system's ship-to state values and create an error report. Furthermore, the script generates an up-loadable sheet to assign the task of correcting the invoice data to a responsible party in Wrike.

The code is organized into the following sections:

    Constants and Variables:
        Constants for specifying the names of the input worksheets and column letters for the relevant data (Sage transaction report and validated Vertex data).
        Worksheet variables to store the references for the input, validated data, error report, and task assignment worksheets.

    Subroutine CompareShipToState():
    The main subroutine that runs the comparison process and generates the error report and task assignment sheet. It performs the following steps:
        Initializes the worksheet references.
        Calls the ClearFormatting subroutine to remove any existing formatting from the input worksheets.
        Loops through each row in the Sage transaction report and compares the ship-to state value against the validated Vertex data. If a discrepancy is found, it adds an entry to the error report.
        Generates a task assignment sheet for correcting the invoice data, which can be uploaded to Wrike.

    Subroutine ClearFormatting(ws As Worksheet):
    A helper subroutine that takes a worksheet as an argument and clears any existing cell formatting (colors, fonts, etc.).

    Functions for retrieving data and row indexes:
        Helper functions to retrieve data from the Sage transaction report and validated Vertex data, and to find row indexes based on search values.

