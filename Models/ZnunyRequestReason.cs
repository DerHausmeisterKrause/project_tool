namespace TaskTool.Models;

public enum ZnunyRequestReason
{
    InitialSync,
    FullTimerSync,
    CandidateTimerSync,
    ManualFullSync,
    ManualCandidateRefresh,
    ManualTicketRefresh,
    ManualDynamicFieldRefresh,
    ManualAssign,
    ManualReply,
    ManualTimeBooking,
    ManualBookingCheck,
    ManualTicketCreate,
    ManualConnectionTest,
    ManualRouteTest
}

public static class ZnunyRequestReasonExtensions
{
    public static bool IsAutomatic(this ZnunyRequestReason reason)
        => reason is ZnunyRequestReason.InitialSync or ZnunyRequestReason.FullTimerSync or ZnunyRequestReason.CandidateTimerSync;
}
