using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using iTextSharp.text.pdf;
using System.IO;
using iTextSharp.text.pdf.security;
using Org.BouncyCastle.Pkcs;

namespace DigitalSign
{
    class Program
    {
        static void Main(string[] args)
        {
            SingleDigitalSignature();

            DigitalSign();

        }

        public static void SingleDigitalSignature()
        {
            // Path to a loadable document.
            string loadPath = @"C:\workspace\PDFDigitalSign\Resource\digitalsignature.docx";
            string savePath = @"C:\workspace\PDFDigitalSign\Resource\Result1.pdf";

            DocumentCore dc = DocumentCore.Load(loadPath);

            // Create a new invisible Shape for the digital signature.       
            // Place the Shape into top-left corner (0 mm, 0 mm) of page.
            Shape signatureShape = new Shape(dc, Layout.Floating(new HorizontalPosition(0f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                                    new VerticalPosition(0f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), new Size(1, 1)));
            ((FloatingLayout)signatureShape.Layout).WrappingStyle = WrappingStyle.InFrontOfText;
            signatureShape.Outline.Fill.SetEmpty();

            // Find a first paragraph and insert our Shape inside it.
            Paragraph firstPar = dc.GetChildElements(true).OfType<Paragraph>().FirstOrDefault();
            firstPar.Inlines.Add(signatureShape);
            
            Picture signaturePict = new Picture(dc, @"C:\workspace\PDFDigitalSign\Resource\sign1.png");

            // Signature picture will be positioned:
            // 14.5 cm from Top of the Shape.
            // 4.5 cm from Left of the Shape.
            signaturePict.Layout = Layout.Floating(
               new HorizontalPosition(4.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
               new VerticalPosition(14.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
               new Size(20, 10, LengthUnit.Millimeter));

            PdfSaveOptions options = new PdfSaveOptions();

            // Path to the certificate (*.pfx).
            options.DigitalSignature.CertificatePath = @"C:\workspace\PDFDigitalSign\Resource\sautinsoft.pfx";
            
            options.DigitalSignature.CertificatePassword = "123456789";
            
            options.DigitalSignature.Location = "World Wide Web";
            options.DigitalSignature.Reason = "Test Signature 1";
            options.DigitalSignature.ContactInfo = "siddhartha.tiwary@outlook.com";
            
            options.DigitalSignature.SignatureLine = signatureShape;
            
            options.DigitalSignature.Signature = signaturePict;

            dc.Save(savePath, options);
            
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
        
        public static void  DigitalSign()
        {
            PdfReader reader = new PdfReader(@"C:\workspace\PDFDigitalSign\Resource\Result1.pdf");
            using (FileStream fout = new FileStream(@"C:\workspace\PDFDigitalSign\Resource\Result2.pdf", FileMode.Create, FileAccess.ReadWrite))
            {
                // appearance
                PdfStamper stamper = PdfStamper.CreateSignature(reader, fout, '\0', null, true);
                PdfSignatureAppearance appearance = stamper.SignatureAppearance;
                //appearance.Reason = SignReason;
                //appearance.Location = SignLocation;
                appearance.SignDate = DateTime.Now.Date;
                appearance.SetVisibleSignature(new iTextSharp.text.Rectangle(100, 100, 50 + 200, 50 + 100), 1, null);//.IsInvisible

                // Custom text and background image
                appearance.Image = iTextSharp.text.Image.GetInstance(@"C:\workspace\PDFDigitalSign\Resource\sign2.png");
                appearance.ImageScale = 0.6f;
                appearance.Image.Alignment = 300;
                appearance.Acro6Layers = true;

                StringBuilder buf = new StringBuilder();
                buf.Append("Digitally Signed by ");
                String name = "Sidd";

                buf.Append(name).Append('\n');
                buf.Append("Date: ").Append(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss zzz"));

                string text = buf.ToString();

                appearance.Layer2Text = text;



                var pk12 = new Pkcs12Store(new System.IO.FileStream(@"C:\workspace\PDFDigitalSign\Resource\certificate.pfx", System.IO.FileMode.Open, System.IO.FileAccess.Read), "12345678".ToCharArray());
                string alias = null;
                foreach (string tAlias in pk12.Aliases)
                {
                    if (pk12.IsKeyEntry(tAlias))
                    {
                        alias = tAlias;
                        break;
                    }
                }
                var pk = pk12.GetKey(alias).Key;

                //digital signature
                IExternalSignature es = new PrivateKeySignature(pk, "SHA-256");

                MakeSignature.SignDetached(appearance, es, new Org.BouncyCastle.X509.X509Certificate[] { pk12.GetCertificate(alias).Certificate }, null, null, null, 0, CryptoStandard.CMS);

                stamper.Close();

            }
        }
    }
}
