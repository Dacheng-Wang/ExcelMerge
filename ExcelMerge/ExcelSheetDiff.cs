﻿using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelSheetDiff
    {
        public ExcelSheet SrcSheet { get; }
        public ExcelSheet DstSheet { get; }
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }

        private bool? hasAnyChanges;
        public bool HasAnyChanges
        {
            get
            {
                if (!hasAnyChanges.HasValue)
                    hasAnyChanges = Rows.Any(r => r.Value.IsModified() || r.Value.IsAdded() || r.Value.IsModified());

                return hasAnyChanges.Value;
            }
        }

        public ExcelSheetDiff(ExcelSheet srcSheet, ExcelSheet dstSheet)
        {
            SrcSheet = srcSheet;
            DstSheet = dstSheet;
            Rows = new SortedDictionary<int, ExcelRowDiff>();
        }

        public ExcelRowDiff CreateRow()
        {
            var row = new ExcelRowDiff(Rows.Any() ? Rows.Keys.Last() + 1 : 0);
            Rows.Add(row.Index, row);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            foreach (var row in Rows)
            {
                if (row.Value.IsAdded())
                    addedRowCount++;
                else if (row.Value.IsRemoved())
                    removedRowCount++;

                if (row.Value.IsModified())
                    modifiedRowCount++;

                modifiedCellCount += row.Value.ModifiedCellCount;
            }

            hasAnyChanges =
                addedRowCount > 0 ||
                removedRowCount > 0 ||
                modifiedRowCount > 0 ||
                modifiedRowCount > 0;

            return new ExcelSheetDiffSummary
            {
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
            };
        }

        public override string ToString()
        {
            return $"{SrcSheet?.Name} {DstSheet?.Name}";
        }
    }
}
