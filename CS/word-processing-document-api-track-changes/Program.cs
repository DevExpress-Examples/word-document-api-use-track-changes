using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Linq;

namespace word_processing_document_api_track_changes
{
    class Program
    {
        static void Main(string[] args)
        {
            RichEditDocumentServer documentProcessor = new RichEditDocumentServer();

            documentProcessor.LoadDocument("DocumentWithRevisions.docx");

            documentProcessor.TrackedMovesConflict += DocumentProcessor_TrackedMovesConflict;
            
            //Turn Track Changes on:
            DocumentTrackChangesOptions documentTrackChangesOptions = documentProcessor.Document.TrackChanges;
            documentTrackChangesOptions.Enabled = true;
            documentTrackChangesOptions.TrackFormatting = true;
            documentTrackChangesOptions.TrackMoves = true;
            
            TrackChangesOptions trackChangesOptions = documentProcessor.Options.Annotations.TrackChanges;

            //Specify how revisions should be displyed in the document:

            trackChangesOptions.DisplayForReviewMode = DisplayForReviewMode.AllMarkup;
            trackChangesOptions.DisplayFormatting = DisplayFormatting.ColorOnly;
            trackChangesOptions.FormattingColor = RevisionColor.ClassicBlue;
            trackChangesOptions.DisplayInsertionStyle = DisplayInsertionStyle.Underline;
            trackChangesOptions.InsertionColor = RevisionColor.DarkRed;

            RevisionCollection documentRevisions = documentProcessor.Document.Revisions;                                  

            //Reject all revisions in the firts page's header:
            SubDocument header = documentProcessor.Document.Sections[0].BeginUpdateHeader(HeaderFooterType.First);
            documentRevisions.RejectAll(header);
            documentProcessor.Document.Sections[0].EndUpdateHeader(header);

            //Reject all revisions from the specific author on the first section:
            var sectionRevisions = documentRevisions.Get(documentProcessor.Document.Sections[0].Range).Where(x => x.Author == "Janet Leverling");
            
            foreach (Revision revision in sectionRevisions)
                    revision.Reject();

            //Accept all format changes:
            documentRevisions.AcceptAll(x => x.Type == RevisionType.CharacterPropertyChanged || x.Type == RevisionType.ParagraphPropertyChanged || x.Type == RevisionType.SectionPropertyChanged);


            documentProcessor.ExportToPdf("DocumentWithAppliedRevisions.pdf");
            System.Diagnostics.Process.Start("DocumentWithAppliedRevisions.pdf");

        }

        private static void DocumentProcessor_TrackedMovesConflict(object sender, TrackedMovesConflictEventArgs e)
        {
            //Compare the length of the original and new location ranges
            //Keep text from the location whose range is the smallest
            e.ResolveMode = (e.OriginalLocationRange.Length <= e.NewLocationRange.Length) ? TrackedMovesConflictResolveMode.KeepOriginalLocationText : TrackedMovesConflictResolveMode.KeepNewLocationText;
        }
    }

}

