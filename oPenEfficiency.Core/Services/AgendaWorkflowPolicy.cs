using System;
using System.IO;

namespace oPenEfficiency.Services
{
    public static class AgendaWorkflowPolicy
    {
        public static bool ShouldCreateChapterInfrastructure(
            int level,
            bool isBreak,
            bool shouldGenerate)
        {
            return shouldGenerate && !isBreak && level == 0;
        }

        public static bool IsSupportedExternalLayoutFile(string path)
        {
            string fileName = Path.GetFileName(path);
            return string.Equals(fileName, "AgendaCreator.xml", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(fileName, "AgendaWizard.xml", StringComparison.OrdinalIgnoreCase);
        }

        public static string GetMutationProtectionError(
            bool presentationProtected,
            bool protectedSlideContentFound,
            bool synchronizeExisting,
            bool createSections)
        {
            if (presentationProtected)
            {
                return "Agenda Creator cannot modify this presentation because it is protected by think-cell.";
            }

            if ((synchronizeExisting || createSections) && protectedSlideContentFound)
            {
                return "Agenda Creator cannot synchronize agenda slides or sections while the presentation contains think-cell-protected slide content.";
            }

            return null;
        }
    }
}
