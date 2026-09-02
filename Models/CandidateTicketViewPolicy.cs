namespace TaskTool.Models;

public static class CandidateTicketViewPolicy
{
    public static IReadOnlyList<ZnunyCandidateTicket> Select(
        IReadOnlyList<ZnunyCandidateTicket> filtered,
        IReadOnlyList<ZnunyCandidateTicket> pool,
        bool showAll)
        => showAll ? pool : filtered;
}
