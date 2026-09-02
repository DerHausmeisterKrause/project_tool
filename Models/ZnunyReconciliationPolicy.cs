namespace TaskTool.Models;

public static class ZnunyReconciliationPolicy
{
    public static bool ShouldCreateTask(bool isCurrentlyAssigned) => isCurrentlyAssigned;
}
