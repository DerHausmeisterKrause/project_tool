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
    my $Names = $Param{Data}->{Names} // '';
    my @Names = grep { $_ ne '' } split /\s*,\s*/, $Names;
    return $Self->{DebuggerObject}->Error( Summary => 'Names is required.' ) if !@Names;

    my $DynamicFieldObject = $Kernel::OM->Get('Kernel::System::DynamicField');
    my $BackendObject      = $Kernel::OM->Get('Kernel::System::DynamicField::Backend');
    my @Fields;

    NAME:
    for my $Name (@Names) {
        my $Config = $DynamicFieldObject->DynamicFieldGet( Name => $Name );
        next NAME if !$Config;
        next NAME if ($Config->{ObjectType} // '') ne 'Ticket';
        next NAME if ($Config->{FieldType}  // '') ne 'Dropdown';

        my $PossibleValues = $BackendObject->PossibleValuesGet(
            DynamicFieldConfig => $Config,
        ) || {};
        my @Options = map {
            +{ Key => "$_", Value => "$PossibleValues->{$_}" }
        } sort keys %{$PossibleValues};

        push @Fields, {
            Name    => $Config->{Name},
            Label   => $Config->{Label} // $Config->{Name},
            Options => \@Options,
        };
    }

    return {
        Success => 1,
        Data    => { Fields => \@Fields },
    };
}

1;
