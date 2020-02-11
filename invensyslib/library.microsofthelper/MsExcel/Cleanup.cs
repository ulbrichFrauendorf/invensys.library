using library.common;
using System;
using Marshal = System.Runtime.InteropServices.Marshal;
namespace library.microsofthelper.MsExcel
{
	public static class Cleanup
	{
		public static void ReleaseObject(object obj)
		{
			try
			{
				if (obj != null && Marshal.IsComObject(obj))
					Marshal.ReleaseComObject(obj);

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
