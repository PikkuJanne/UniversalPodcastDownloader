@{
    # The existing interactive TUI and developer command summaries intentionally
    # use host output on PowerShell 5.1+. All other default rules remain enabled.
    ExcludeRules = @('PSAvoidUsingWriteHost')
}
