@{
    TargetDir = 'C:\Meine Programme\Kindlimann'
    OutputFileName = 'Kindlimann Bestellliste.xlsm'
    LogoPath = 'C:\Meine Programme\Metallraum Logo.jpg'
    Password = '9500Wil'
    MinimumStartDate = '2026-06-29'

    ProfileList = @(
        '40 x 15 x 2 mm'
        'FL 30 x 6 mm'
        'FL 40 x 10 mm'
        'FL 40 x 15 mm'
        'RD 10 mm'
    )

    Mail = @{
        To = 'claudia.desantis@kindlimann.ch'
        Salutation = 'Guten Tag Frau De Santis'
        SubjectPrefix = 'Sammelbestellung Metallraum'
        Closing = 'Freundliche Grüsse'
        Signature = 'Metallraum'
    }

    Layout = @{
        ButtonGapPoints = 24
        PdfMarginCm = 2
        Columns = @{
            A = 25
            B = 60
            C = 20
            D = 48
        }
        Rows = @{
            Header = 64
            Meta = 28
            Status = 22
            Spacer = 16
            OrderTitle = 28
            OrderInfo = 32
            ProfileHeader = 24
            Profile = 27
            Blank = 18
        }
    }
}
