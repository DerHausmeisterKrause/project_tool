using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class PlenaroShareTests
{
    private const string SharedTaskId = "11111111-1111-1111-1111-111111111111";
    private const string FirstSegmentId = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string SecondSegmentId = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";

    [Fact]
    public void AttendeesAreValidatedAndDeduplicatedCaseInsensitively()
    {
        Assert.True(SegmentAttendeeParser.TryParse("Max@firma.de; max@FIRMA.de, anna@firma.de\nbob@firma.de", out var values, out _));
        Assert.Equal(3, values.Count);
        Assert.False(SegmentAttendeeParser.TryParse("valid@firma.de; xyz", out _, out var invalid));
        Assert.Equal("xyz", invalid);
    }

    [Fact]
    public void PayloadRoundTripsAndTamperingIsRejected()
    {
        var payload = Sample();
        var body = PlenaroShareCodec.AppendOrReplace("Beschreibung", payload, out var hash);
        Assert.True(PlenaroShareCodec.TryParse(body, out var parsed, out var parsedHash, out _));
        Assert.Equal(payload.TaskShareId, parsed!.TaskShareId);
        Assert.Equal(hash, parsedHash);
        Assert.False(PlenaroShareCodec.TryParse(body.Replace("Payload: ", "Payload: A"), out _, out _, out var reason));
        Assert.Contains(reason, new[] { "checksum-invalid", "json-invalid", "base64-invalid" });
    }

    [Fact]
    public void ExistingBlockIsReplacedAndHumanBodyRemains()
    {
        var first = PlenaroShareCodec.AppendOrReplace("Menschlich", Sample(), out _);
        var second = PlenaroShareCodec.AppendOrReplace(first, Sample(), out _);
        Assert.StartsWith("Menschlich", second);
        Assert.Equal(1, second.Split(PlenaroShareCodec.BeginMarker).Length - 1);
    }

    [Fact]
    public void UnknownVersionAndAmbiguousBlocksAreRejected()
    {
        var body = PlenaroShareCodec.AppendOrReplace("Text", Sample(), out _);
        Assert.False(PlenaroShareCodec.TryParse(body.Replace("Version: 1", "Version: 2"), out _, out _, out _));
        Assert.False(PlenaroShareCodec.TryParse(body + "\n" + body, out _, out _, out var reason));
        Assert.Equal("ambiguous-blocks", reason);
    }

    [Fact]
    public void V29ToV30MigrationIsAdditiveAndKeepsData()
    {
        var path=Path.GetTempFileName();
        try { using(var c=new SqliteConnection($"Data Source={path}")){c.Open();using var cmd=c.CreateCommand();cmd.CommandText="CREATE TABLE schema_version(version INTEGER NOT NULL); INSERT INTO schema_version VALUES(29); CREATE TABLE tasks(id TEXT PRIMARY KEY,title TEXT NOT NULL,description TEXT,ticket_url TEXT,start_local TEXT,end_local TEXT,status TEXT NOT NULL,priority INTEGER,tags TEXT,outlook_entry_id TEXT,ticket_minutes_booked INTEGER NOT NULL DEFAULT 0,ticket_seconds_booked INTEGER NOT NULL DEFAULT 0,is_pinned INTEGER NOT NULL DEFAULT 0,created_utc TEXT NOT NULL,updated_utc TEXT NOT NULL); INSERT INTO tasks VALUES('keep','Keep','','',NULL,NULL,'Planned',NULL,'','',0,0,0,'2020-01-01','2020-01-01'); CREATE TABLE plenaro_shared_task_imports(task_share_id TEXT PRIMARY KEY,local_task_id TEXT NOT NULL,origin_client_instance_id TEXT NOT NULL DEFAULT '',last_payload_hash TEXT NOT NULL DEFAULT ''); INSERT INTO plenaro_shared_task_imports VALUES('share','keep','origin','hash');";cmd.ExecuteNonQuery();}
            new DatabaseService(new LoggerService(AppLogLevel.Error),path).Initialize();
            using var check=new SqliteConnection($"Data Source={path}");check.Open();using var cmd2=check.CreateCommand();cmd2.CommandText="SELECT (SELECT version FROM schema_version),(SELECT title FROM tasks WHERE id='keep'),(SELECT COUNT(*) FROM pragma_table_info('plenaro_shared_task_imports') WHERE name='last_generated_utc'),(SELECT last_payload_hash FROM plenaro_shared_task_imports WHERE task_share_id='share')";using var r=cmd2.ExecuteReader();Assert.True(r.Read());Assert.Equal(30,r.GetInt32(0));Assert.Equal("Keep",r.GetString(1));Assert.Equal(1,r.GetInt32(2));Assert.Equal("hash",r.GetString(3));
        } finally { SqliteConnection.ClearAllPools(); File.Delete(path); }
    }

    [Fact]
    public void SharedUnassignedZnunyTaskRemainsVisible()
    { var task=new TaskItem{Tags="ZnunyTicketID:42",IsZnunyAssigned=false,IsPlenaroShared=true};Assert.True(task.IsOperationallyVisible); }

    [Fact]
    public void MeetingRemovalDecisionOnlyCancelsAnExistingMeetingWithRecipients()
    {
        Assert.True(OutlookInteropService.ShouldCancelMeeting(1, 2, 0));
        Assert.False(OutlookInteropService.ShouldCancelMeeting(0, 2, 0));
        Assert.False(OutlookInteropService.ShouldCancelMeeting(1, 0, 0));
        Assert.False(OutlookInteropService.ShouldCancelMeeting(1, 2, 1));
    }

    [Fact]
    public void AccessDeniedIsDetectedStructurally()
    {
        Assert.True(ViewModels.TodayViewModel.IsZnunyAccessDenied(new ZnunyApiException("TicketGet", System.Net.HttpStatusCode.Forbidden, "Whatever", "", "")));
        Assert.True(ViewModels.TodayViewModel.IsZnunyAccessDenied(new ZnunyApiException("TicketGet", System.Net.HttpStatusCode.BadRequest, "AccessDenied", "", "")));
        Assert.False(ViewModels.TodayViewModel.IsZnunyAccessDenied(new ZnunyApiException("TicketGet", System.Net.HttpStatusCode.BadGateway, "Protocol", "Access denied text is irrelevant", "")));
    }

    [Fact]
    public void ShareRangeIncludesTodayPlus180AndUsesExclusiveEnd()
    {
        using var fixture = new ImportFixture();
        using var coordinator = new PlenaroShareImportCoordinator(fixture.Calendar, fixture.Importer, fixture.Settings, fixture.Logger,
            () => new DateTime(2026, 9, 14));
        var range = coordinator.GetRange();
        Assert.Equal(new DateTime(2026, 8, 15), range.FromInclusive);
        Assert.Equal(new DateTime(2027, 3, 14), range.ToExclusive);
    }

    [Fact]
    public void MultipleSegmentsAreIdempotentAndSnapshotsNeverMoveBackward()
    {
        using var fixture = new ImportFixture();
        var generated = new DateTime(2026, 9, 14, 12, 0, 0, DateTimeKind.Utc);
        var first = fixture.Event(Payload(SharedTaskId, FirstSegmentId, "New title", generated, 9));
        var second = fixture.Event(Payload(SharedTaskId, SecondSegmentId, "New title", generated, 11));
        fixture.Importer.Import([first, second]);

        var task = Assert.Single(fixture.Tasks.GetAllTasks());
        Assert.Equal(2, fixture.Tasks.GetSegments(task.Id).Count);
        var updated = task.UpdatedUtc;
        fixture.Importer.Import([first, second]);
        task = Assert.Single(fixture.Tasks.GetAllTasks());
        Assert.Equal(updated, task.UpdatedUtc);
        Assert.Equal(2, fixture.Tasks.GetSegments(task.Id).Count);

        fixture.Importer.Import([fixture.Event(Payload(SharedTaskId, FirstSegmentId, "Old title", generated.AddMinutes(-1), 10))]);
        task = Assert.Single(fixture.Tasks.GetAllTasks());
        Assert.Equal("New title", task.Title);
        var segments = fixture.Tasks.GetSegments(task.Id);
        Assert.Equal(10, segments.Single(segment => segment.SegmentShareId == FirstSegmentId).StartLocal.Hour);
        Assert.Equal(11, segments.Single(segment => segment.SegmentShareId == SecondSegmentId).StartLocal.Hour);

        fixture.Importer.Import([fixture.Event(Payload(SharedTaskId, FirstSegmentId, "Newest title", generated.AddMinutes(1), 12))]);
        Assert.Equal("Newest title", Assert.Single(fixture.Tasks.GetAllTasks()).Title);
        Assert.Equal(2, fixture.Tasks.GetSegments(task.Id).Count);
    }

    [Fact]
    public void MissingSegmentMappingAndMissingTaskMappingRecoverWithoutDuplicateTask()
    {
        using var fixture = new ImportFixture();
        var payload = Payload(SharedTaskId, FirstSegmentId, "Recovery", DateTime.UtcNow, 9);
        fixture.Importer.Import([fixture.Event(payload)]);
        var task = Assert.Single(fixture.Tasks.GetAllTasks());
        var segment = Assert.Single(fixture.Tasks.GetSegments(task.Id));
        using (var connection = new SqliteConnection(fixture.Database.ConnectionString))
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM plenaro_shared_segment_imports; DELETE FROM task_segments;";
            command.ExecuteNonQuery();
        }
        fixture.Importer.Import([fixture.Event(payload)]);
        Assert.Single(fixture.Tasks.GetAllTasks());
        Assert.Single(fixture.Tasks.GetSegments(task.Id));

        using (var connection = new SqliteConnection(fixture.Database.ConnectionString))
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM plenaro_shared_task_imports; DELETE FROM plenaro_shared_segment_imports; DELETE FROM task_segments;";
            command.ExecuteNonQuery();
        }
        payload.SegmentShareId = SecondSegmentId;
        fixture.Importer.Import([fixture.Event(payload)]);
        Assert.Single(fixture.Tasks.GetAllTasks());
        Assert.Single(fixture.Tasks.GetSegments(task.Id));
        Assert.NotEqual(segment.Id, fixture.Tasks.GetSegments(task.Id)[0].Id);
    }

    [Fact]
    public void MissingSegmentMappingReusesSegmentShareIdAndRepairsMapping()
    {
        using var fixture = new ImportFixture();
        var payload = Payload(SharedTaskId, FirstSegmentId, "Recovery", DateTime.UtcNow, 9);
        fixture.Importer.Import([fixture.Event(payload)]);
        var task = Assert.Single(fixture.Tasks.GetAllTasks());
        var originalSegment = Assert.Single(fixture.Tasks.GetSegments(task.Id));

        using (var connection = new SqliteConnection(fixture.Database.ConnectionString))
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM plenaro_shared_segment_imports;";
            command.ExecuteNonQuery();
        }

        fixture.Importer.Import([fixture.Event(payload)]);

        Assert.Single(fixture.Tasks.GetAllTasks());
        var recoveredSegment = Assert.Single(fixture.Tasks.GetSegments(task.Id));
        Assert.Equal(originalSegment.Id, recoveredSegment.Id);
        using var check = new SqliteConnection(fixture.Database.ConnectionString);
        check.Open();
        using var checkCommand = check.CreateCommand();
        checkCommand.CommandText = "SELECT local_segment_id FROM plenaro_shared_segment_imports WHERE segment_share_id=$id";
        checkCommand.Parameters.AddWithValue("$id", FirstSegmentId);
        Assert.Equal(originalSegment.Id, (long)checkCommand.ExecuteScalar()!);
    }

    private static PlenaroSharePayloadV1 Sample()=>new(){TaskShareId=Guid.NewGuid().ToString(),SegmentShareId=Guid.NewGuid().ToString(),OriginClientInstanceId=Guid.NewGuid().ToString(),TaskTitle="Titel",TaskDescription="Beschreibung",SegmentStartLocal=DateTime.Today.AddHours(9),SegmentEndLocal=DateTime.Today.AddHours(10),GeneratedUtc=DateTime.UtcNow};

    private static PlenaroSharePayloadV1 Payload(string taskId, string segmentId, string title, DateTime generated, int hour) => new()
    {
        TaskShareId = taskId, SegmentShareId = segmentId, OriginClientInstanceId = "remote-client", TaskTitle = title,
        TaskDescription = "Description", SegmentNote = segmentId, SegmentStartLocal = DateTime.Today.AddHours(hour),
        SegmentEndLocal = DateTime.Today.AddHours(hour + 1), GeneratedUtc = generated
    };

    private sealed class ImportFixture : IDisposable
    {
        private readonly string _databasePath = Path.GetTempFileName();
        private readonly string _settingsPath = Path.GetTempFileName();
        public LoggerService Logger { get; } = new(AppLogLevel.Error);
        public DatabaseService Database { get; }
        public SettingsService Settings { get; }
        public TaskService Tasks { get; }
        public PlenaroShareImportService Importer { get; }
        public OutlookCalendarService Calendar { get; }

        public ImportFixture()
        {
            File.Delete(_settingsPath);
            Settings = new SettingsService(Logger, _settingsPath);
            Settings.Current.OutlookCalendarEnabled = false;
            Database = new DatabaseService(Logger, _databasePath);
            Database.Initialize();
            var outlook = new OutlookInteropService(Logger, Settings);
            Tasks = new TaskService(Database, Logger, outlook, Settings);
            Importer = new PlenaroShareImportService(Database, Tasks, Settings, Logger);
            Calendar = new OutlookCalendarService(Logger, Settings, outlook, new WorkDayService(Database, Logger));
        }

        public OutlookCalendarEvent Event(PlenaroSharePayloadV1 payload) => new()
        {
            StartLocal = payload.SegmentStartLocal,
            EndLocal = payload.SegmentEndLocal,
            FullBody = PlenaroShareCodec.AppendOrReplace(string.Empty, payload, out _)
        };

        public void Dispose()
        {
            Calendar.Dispose();
            SqliteConnection.ClearAllPools();
            File.Delete(_databasePath);
            File.Delete(_settingsPath);
        }
    }
}
