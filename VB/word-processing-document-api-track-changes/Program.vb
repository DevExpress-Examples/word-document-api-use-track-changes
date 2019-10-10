Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Linq

Namespace word_processing_document_api_track_changes
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim documentProcessor As New RichEditDocumentServer()

			documentProcessor.LoadDocument("DocumentWithRevisions.docx")

			AddHandler documentProcessor.TrackedMovesConflict, AddressOf DocumentProcessor_TrackedMovesConflict

			'Turn Track Changes on:
			Dim documentTrackChangesOptions As DocumentTrackChangesOptions = documentProcessor.Document.TrackChanges
			documentTrackChangesOptions.Enabled = True
			documentTrackChangesOptions.TrackFormatting = True
			documentTrackChangesOptions.TrackMoves = True

			Dim trackChangesOptions As TrackChangesOptions = documentProcessor.Options.Annotations.TrackChanges

			'Specify how revisions should be displyed in the document:

			trackChangesOptions.DisplayForReviewMode = DisplayForReviewMode.AllMarkup
			trackChangesOptions.DisplayFormatting = DisplayFormatting.ColorOnly
			trackChangesOptions.FormattingColor = RevisionColor.ClassicBlue
			trackChangesOptions.DisplayInsertionStyle = DisplayInsertionStyle.Underline
			trackChangesOptions.InsertionColor = RevisionColor.DarkRed

			Dim documentRevisions As RevisionCollection = documentProcessor.Document.Revisions

			'Reject all revisions in the firts page's header:
			Dim header As SubDocument = documentProcessor.Document.Sections(0).BeginUpdateHeader(HeaderFooterType.First)
			documentRevisions.RejectAll(header)
			documentProcessor.Document.Sections(0).EndUpdateHeader(header)

			'Reject all revisions from the specific author on the first section:
			Dim sectionRevisions = documentRevisions.Get(documentProcessor.Document.Sections(0).Range).Where(Function(x) x.Author = "Janet Leverling")

			For Each revision As Revision In sectionRevisions
					revision.Reject()
			Next revision

			'Accept all format changes:
			documentRevisions.AcceptAll(Function(x) x.Type = RevisionType.CharacterPropertyChanged OrElse x.Type = RevisionType.ParagraphPropertyChanged OrElse x.Type = RevisionType.SectionPropertyChanged)


			documentProcessor.ExportToPdf("DocumentWithAppliedRevisions.pdf")
			System.Diagnostics.Process.Start("DocumentWithAppliedRevisions.pdf")

		End Sub

		Private Shared Sub DocumentProcessor_TrackedMovesConflict(ByVal sender As Object, ByVal e As TrackedMovesConflictEventArgs)
			'Compare the length of the original and new location ranges
			'Keep text from the location whose range is the smallest
			e.ResolveMode = If(e.OriginalLocationRange.Length <= e.NewLocationRange.Length, TrackedMovesConflictResolveMode.KeepOriginalLocationText, TrackedMovesConflictResolveMode.KeepNewLocationText)
		End Sub
	End Class

End Namespace

