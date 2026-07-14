using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grid.Office
{
    internal sealed class OfficeAutomationService
    {
        private readonly ExcelAutomationService _excel;
        private readonly WordAutomationService _word;
        private readonly PowerPointAutomationService _powerPoint;

        public OfficeAutomationService(Excel.Application application, OfficeDispatcher dispatcher)
        {
            if (application == null)
            {
                throw new ArgumentNullException(nameof(application));
            }

            if (dispatcher == null)
            {
                throw new ArgumentNullException(nameof(dispatcher));
            }

            _excel = new ExcelAutomationService(application, dispatcher);
            _word = new WordAutomationService(dispatcher);
            _powerPoint = new PowerPointAutomationService(dispatcher);
        }

        public ExcelAutomationService Excel
        {
            get { return _excel; }
        }

        public WordAutomationService Word
        {
            get { return _word; }
        }

        public PowerPointAutomationService PowerPoint
        {
            get { return _powerPoint; }
        }

        public async Task<Dictionary<string, object>> ListRunningAppsAsync(CancellationToken cancellationToken)
        {
            Dictionary<string, object> wordContext;
            Dictionary<string, object> powerPointContext;

            wordContext = await _word.GetContextAsync(cancellationToken).ConfigureAwait(false);
            powerPointContext = await _powerPoint.GetContextAsync(cancellationToken).ConfigureAwait(false);

            return new Dictionary<string, object>
            {
                ["excel"] = true,
                ["word"] = wordContext.ContainsKey("running") && (bool)wordContext["running"],
                ["powerpoint"] = powerPointContext.ContainsKey("running") && (bool)powerPointContext["running"]
            };
        }

        public async Task<Dictionary<string, object>> GetActiveContextAsync(CancellationToken cancellationToken)
        {
            Dictionary<string, object> excelContext;
            Dictionary<string, object> wordContext;
            Dictionary<string, object> powerPointContext;

            excelContext = await _excel.GetContextAsync(cancellationToken).ConfigureAwait(false);
            wordContext = await _word.GetContextAsync(cancellationToken).ConfigureAwait(false);
            powerPointContext = await _powerPoint.GetContextAsync(cancellationToken).ConfigureAwait(false);

            return new Dictionary<string, object>
            {
                ["excel"] = excelContext,
                ["word"] = wordContext,
                ["powerpoint"] = powerPointContext
            };
        }
    }
}
