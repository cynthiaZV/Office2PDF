using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSPDF {
	interface IMSConvert {

		/// <summary>
		/// 
		/// </summary>
		/// <returns>True if conversion succecss, False otherwise</returns>
		Boolean Convert();
		/// <summary>
		///		Close MS Application. Throws ApplicationException upon failure.  
		/// </summary>
		void Close();
	}
}
