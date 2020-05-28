using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	public partial class SystemMail {
		private static string mailUser = Properties.Settings.Default.MailUser;
		private static string mailPassword = Properties.Settings.Default.MailPassword;
		private static string mailDomain = Properties.Settings.Default.MailDomain;
		private static string mailSmtpServer = Properties.Settings.Default.MailSmtpServer;
	}
}
