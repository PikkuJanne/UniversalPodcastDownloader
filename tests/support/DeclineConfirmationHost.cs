// Test-only host: exercises the real ShouldProcess prompt with a deterministic No.
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Security;

namespace UpdTests
{
    public sealed class DeclineConfirmationHost : PSHost
    {
        private readonly Guid instanceId = Guid.NewGuid();
        private readonly DeclineConfirmationUI userInterface = new DeclineConfirmationUI();
        public int ConfirmationCount { get { return userInterface.ConfirmationCount; } }
        public override Guid InstanceId { get { return instanceId; } }
        public override string Name { get { return "UPD owned confirmation fixture"; } }
        public override Version Version { get { return new Version(1, 0); } }
        public override PSHostUserInterface UI { get { return userInterface; } }
        public override CultureInfo CurrentCulture { get { return CultureInfo.InvariantCulture; } }
        public override CultureInfo CurrentUICulture { get { return CultureInfo.InvariantCulture; } }
        public override void EnterNestedPrompt() { throw new InvalidOperationException("Unexpected nested prompt."); }
        public override void ExitNestedPrompt() { throw new InvalidOperationException("Unexpected nested prompt exit."); }
        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }
        public override void SetShouldExit(int exitCode) { throw new InvalidOperationException("Unexpected process exit."); }
    }

    public sealed class DeclineConfirmationUI : PSHostUserInterface
    {
        public int ConfirmationCount { get; private set; }
        public override PSHostRawUserInterface RawUI { get { return null; } }
        public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice)
        {
            ConfirmationCount++;
            for (int index = 0; index < choices.Count; index++)
                if (choices[index].Label.Replace("&", "") == "No") return index;
            throw new InvalidOperationException("No explicit No choice was offered.");
        }
        public override Dictionary<string, PSObject> Prompt(string caption, string message, Collection<FieldDescription> descriptions)
        { throw new InvalidOperationException("Unexpected field prompt."); }
        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName)
        { throw new InvalidOperationException("Unexpected credential prompt."); }
        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes types, PSCredentialUIOptions options)
        { throw new InvalidOperationException("Unexpected credential prompt."); }
        public override string ReadLine() { throw new InvalidOperationException("Unexpected line prompt."); }
        public override SecureString ReadLineAsSecureString() { throw new InvalidOperationException("Unexpected secure prompt."); }
        public override void Write(string value) { }
        public override void Write(ConsoleColor foreground, ConsoleColor background, string value) { }
        public override void WriteLine(string value) { }
        public override void WriteDebugLine(string message) { }
        public override void WriteErrorLine(string value) { }
        public override void WriteProgress(long sourceId, ProgressRecord record) { }
        public override void WriteVerboseLine(string message) { }
        public override void WriteWarningLine(string message) { }
    }
}
