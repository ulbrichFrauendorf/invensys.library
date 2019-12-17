using library.common;
using System;
using System.Collections.Generic;
using System.Text;

namespace library.microsofthelper.MsExcel
{
	public static class Cleanup
	{
		public static void ReleaseObject(object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				throw new LocalSystemException("Com Object not Released.", ex);
			}
			finally
			{
				GC.Collect();
			}
		}
	}
}
