package Kernel::System::GenericInterface::Operation::DynamicField::Options;

use strict;
use warnings;

our $ObjectManagerDisabled = 1;

sub new {
    my ( $Type, %Param ) = @_;
    my $Self = bless {}, $Type;
    for my $Required (qw(DebuggerObject WebserviceID)) {
        return if !$Param{$Required};
        $Self->{$Required} = $Param{$Required};
    }
    return $Self;
}

sub Run {
    my ( $Self, %Param ) = @_;
    my $Name = $Param{Data}->{FieldName} // '';
    return $Self->{DebuggerObject}->Error( Summary => 'FieldName is required.' ) if $Name eq '';

    my $DynamicFieldObject = $Kernel::OM->Get('Kernel::System::DynamicField');
    my $BackendObject      = $Kernel::OM->Get('Kernel::System::DynamicField::Backend');
    my $Config = $DynamicFieldObject->DynamicFieldGet( Name => $Name );
    return $Self->{DebuggerObject}->Error( Summary => 'Field is not an allowed ticket dropdown.' )
        if !$Config
        || ($Config->{ObjectType} // '') ne 'Ticket'
        || ($Config->{FieldType}  // '') ne 'Dropdown';

    my $PossibleValues = $BackendObject->PossibleValuesGet(
        DynamicFieldConfig => $Config,
    ) || {};
    my @Options = map {
        +{ Key => "$_", Value => "$PossibleValues->{$_}" }
    } sort keys %{$PossibleValues};

    return {
        Success => 1,
        Data    => {
            Field => {
                Name    => $Config->{Name},
                Label   => $Config->{Label} // $Config->{Name},
                Options => \@Options,
            },
        },
    };
}

1;
