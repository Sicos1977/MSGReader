// -- FILE ------------------------------------------------------------------
// name       : RtfDocumentInfo.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfDocumentInfo : IRtfDocumentInfo
	{

		// ----------------------------------------------------------------------
		public int? Id
		{
			get { return id; }
			set { id = value; }
		} // Id

		// ----------------------------------------------------------------------
		public int? Version
		{
			get { return version; }
			set { version = value; }
		} // Version

		// ----------------------------------------------------------------------
		public int? Revision
		{
			get { return revision; }
			set { revision = value; }
		} // Revision

		// ----------------------------------------------------------------------
		public string Title
		{
			get { return title; }
			set { title = value; }
		} // Title

		// ----------------------------------------------------------------------
		public string Subject
		{
			get { return subject; }
			set { subject = value; }
		} // Subject

		// ----------------------------------------------------------------------
		public string Author
		{
			get { return author; }
			set { author = value; }
		} // Author

		// ----------------------------------------------------------------------
		public string Manager
		{
			get { return manager; }
			set { manager = value; }
		} // Manager

		// ----------------------------------------------------------------------
		public string Company
		{
			get { return company; }
			set { company = value; }
		} // Company

		// ----------------------------------------------------------------------
		public string Operator
		{
			get { return operatorName; }
			set { operatorName = value; }
		} // Operator

		// ----------------------------------------------------------------------
		public string Category
		{
			get { return category; }
			set { category = value; }
		} // Category

		// ----------------------------------------------------------------------
		public string Keywords
		{
			get { return keywords; }
			set { keywords = value; }
		} // Keywords

		// ----------------------------------------------------------------------
		public string Comment
		{
			get { return comment; }
			set { comment = value; }
		} // Comment

		// ----------------------------------------------------------------------
		public string DocumentComment
		{
			get { return documentComment; }
			set { documentComment = value; }
		} // DocumentComment

		// ----------------------------------------------------------------------
		public string HyperLinkbase
		{
			get { return hyperLinkbase; }
			set { hyperLinkbase = value; }
		} // HyperLinkbase

		// ----------------------------------------------------------------------
		public DateTime? CreationTime
		{
			get { return creationTime; }
			set { creationTime = value; }
		} // CreationTime

		// ----------------------------------------------------------------------
		public DateTime? RevisionTime
		{
			get { return revisionTime; }
			set { revisionTime = value; }
		} // RevisionTime

		// ----------------------------------------------------------------------
		public DateTime? PrintTime
		{
			get { return printTime; }
			set { printTime = value; }
		} // PrintTime

		// ----------------------------------------------------------------------
		public DateTime? BackupTime
		{
			get { return backupTime; }
			set { backupTime = value; }
		} // BackupTime

		// ----------------------------------------------------------------------
		public int? NumberOfPages
		{
			get { return numberOfPages; }
			set { numberOfPages = value; }
		} // NumberOfPages

		// ----------------------------------------------------------------------
		public int? NumberOfWords
		{
			get { return numberOfWords; }
			set { numberOfWords = value; }
		} // NumberOfWords

		// ----------------------------------------------------------------------
		public int? NumberOfCharacters
		{
			get { return numberOfCharacters; }
			set { numberOfCharacters = value; }
		} // NumberOfCharacters

		// ----------------------------------------------------------------------
		public int? EditingTimeInMinutes
		{
			get { return editingTimeInMinutes; }
			set { editingTimeInMinutes = value; }
		} // EditingTimeInMinutes

		// ----------------------------------------------------------------------
		public void Reset()
		{
			id = null;
			version = null;
			revision = null;
			title = null;
			subject = null;
			author = null;
			manager = null;
			company = null;
			operatorName = null;
			category = null;
			keywords = null;
			comment = null;
			documentComment = null;
			hyperLinkbase = null;
			creationTime = null;
			revisionTime = null;
			printTime = null;
			backupTime = null;
			numberOfPages = null;
			numberOfWords = null;
			numberOfCharacters = null;
			editingTimeInMinutes = null;
		} // Reset

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return "RTFDocInfo";
		} // ToString

		// ----------------------------------------------------------------------
		// members
		private int? id;
		private int? version;
		private int? revision;
		private string title;
		private string subject;
		private string author;
		private string manager;
		private string company;
		private string operatorName;
		private string category;
		private string keywords;
		private string comment;
		private string documentComment;
		private string hyperLinkbase;
		private DateTime? creationTime;
		private DateTime? revisionTime;
		private DateTime? printTime;
		private DateTime? backupTime;
		private int? numberOfPages;
		private int? numberOfWords;
		private int? numberOfCharacters;
		private int? editingTimeInMinutes;

	} // interface IRtfDocumentInfo

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
