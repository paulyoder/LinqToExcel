using System.Collections.Generic;

namespace LinqToExcel
{
	/// <summary>
	/// Implement this interface to receive values for cells that
	/// were not mapped.
	/// </summary>
	public interface IContainsUnmappedCells
	{
		IDictionary<string, Cell> UnmappedCells { get; }
	}
}
