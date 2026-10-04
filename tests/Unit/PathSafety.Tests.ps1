BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/PathSafety.ps1')
    Mock Invoke-WebRequest { throw 'Path tests must not make network requests.' }
}

Describe 'A009: canonical Windows containment helpers' -Tag 'Unit', 'A009' {
    It 'normalizes a normal absolute drive root without accessing or creating it' {
        Get-PodcastCanonicalRoot -Path 'C:/Synthetic/Podcasts/' | Should -Be 'C:\Synthetic\Podcasts'
        Get-PodcastCanonicalRoot -Path 'C:\' | Should -Be 'C:\'
        Get-PodcastCanonicalRoot -Path '\\synthetic.invalid\share\Podcasts' |
            Should -Be '\\synthetic.invalid\share\Podcasts'
    }

    It 'rejects ambiguous or device roots <Root>' -ForEach @(
        @{ Root = '' }, @{ Root = 'relative\podcasts' }, @{ Root = 'C:podcasts' },
        @{ Root = '\podcasts' }, @{ Root = '\\?\C:\podcasts' }, @{ Root = '\\.\C:\podcasts' },
        @{ Root = '\\?\UNC\host\share' }, @{ Root = 'FileSystem::C:\podcasts' },
        @{ Root = 'C:\safe\..\outside' }, @{ Root = 'C:\safe\.\podcasts' },
        @{ Root = 'C:\safe.\podcasts' }, @{ Root = 'C:\safe \podcasts' },
        @{ Root = 'C:\safe\\podcasts' }, @{ Root = '\\host' }, @{ Root = 'C:\\' }
    ) {
        { Get-PodcastCanonicalRoot -Path $Root } | Should -Throw
    }

    It 'preserves a safe filename within the root' {
        Get-PodcastDestination -Root 'C:\Synthetic\Podcasts' -RelativePath 'Show/episode.mp3' |
            Should -Be 'C:\Synthetic\Podcasts\Show\episode.mp3'
    }

    It 'rejects hostile relative destinations <Path>' -ForEach @(
        @{ Path = '' }, @{ Path = '.' }, @{ Path = '..' }, @{ Path = '..\sibling\episode.mp3' },
        @{ Path = 'Show\..\episode.mp3' }, @{ Path = 'Show\.\episode.mp3' },
        @{ Path = 'C:\outside.mp3' }, @{ Path = 'C:outside.mp3' }, @{ Path = '\outside.mp3' },
        @{ Path = '\\host\share\outside.mp3' }, @{ Path = '\\?\C:\outside.mp3' },
        @{ Path = 'Show\\episode.mp3' }, @{ Path = 'Show\' },
        @{ Path = 'Show\episode.mp3:stream' }, @{ Path = 'Show\episode?.mp3' },
        @{ Path = 'Show\episode.mp3.' }, @{ Path = 'Show\episode.mp3 ' },
        @{ Path = 'CON' }, @{ Path = 'nul.mp3' }, @{ Path = 'CoM1.mp3' },
        @{ Path = 'LPT9.log' }, @{ Path = 'CONIN$.mp3' }, @{ Path = 'CON .mp3' }
    ) {
        { Get-PodcastDestination -Root 'C:\Synthetic\Podcasts' -RelativePath $Path } | Should -Throw
    }

    It 'rejects control characters and superscript device names' {
        { Get-PodcastDestination -Root 'C:\Synthetic' -RelativePath ("bad$([char]0)name.mp3") } | Should -Throw
        { Get-PodcastDestination -Root 'C:\Synthetic' -RelativePath ("LPT$([char]0x00b2).mp3") } | Should -Throw
    }

    It 'enforces file and directory budgets at their boundary' {
        $root = 'C:\' + ('r' * 100)
        (Get-PodcastDestination -Root $root -RelativePath ('f' * 155)).Length | Should -Be 259
        { Get-PodcastDestination -Root $root -RelativePath ('f' * 156) } | Should -Throw '*limit of 259*'
        (Get-PodcastDestination -Root $root -RelativePath ('f' * 143) -Directory).Length | Should -Be 247
        { Get-PodcastDestination -Root $root -RelativePath ('f' * 144) -Directory } | Should -Throw '*limit of 247*'
        { Get-PodcastCanonicalRoot -Path ('C:\' + ('r' * 245)) } | Should -Throw '*limit of 247*'
        { Get-PodcastDestination -Root 'C:\' -RelativePath ('f' * 256) } | Should -Throw '*255*'
    }
}

Describe 'A009: destination disk checks' -Tag 'Unit', 'A009' {
    It 'checks a future destination without creating directories or files' {
        $root = Join-Path $TestDrive 'missing-root'
        $result = Assert-PodcastDestination -Root $root -RelativePath 'Show\episode.mp3'
        $result | Should -Be (Join-Path $root 'Show\episode.mp3')
        Test-Path -LiteralPath $root | Should -BeFalse
    }

    It 'accepts an ordinary existing target without altering its bytes' {
        $file = Join-Path $TestDrive 'unchanged.mp3'
        [IO.File]::WriteAllText($file, 'owned unit fixture')
        Assert-PodcastDestination -Root $TestDrive -RelativePath 'unchanged.mp3' | Should -Be $file
        [IO.File]::ReadAllText($file) | Should -Be 'owned unit fixture'
    }

    It 'rejects a file used as the root, a directory destination, or an ancestor' {
        $file = Join-Path $TestDrive 'file-parent'
        [IO.File]::WriteAllText($file, 'owned unit fixture')
        { Assert-PodcastDestination -Root $file } | Should -Throw '*existing file*'
        { Assert-PodcastDestination -Root $TestDrive -RelativePath 'file-parent' -Directory } | Should -Throw '*existing file*'
        { Assert-PodcastDestination -Root $TestDrive -RelativePath 'file-parent\episode.mp3' } | Should -Throw '*existing file*'
    }

    It 'rejects a directory used as a media destination' {
        $directory = Join-Path $TestDrive 'directory.mp3'
        $null = [IO.Directory]::CreateDirectory($directory)
        { Assert-PodcastDestination -Root $TestDrive -RelativePath 'directory.mp3' } | Should -Throw '*existing directory*'
    }

    It 'rejects a differently cased sibling even if exact-path lookup reports it absent' {
        $file = Join-Path $TestDrive 'CASE.mp3'
        [IO.File]::WriteAllText($file, 'owned unit fixture')
        Mock Get-PodcastPathAttribute { return $null } -ParameterFilter { $LiteralPath -ceq (Join-Path $TestDrive 'case.mp3') }
        { Assert-PodcastDestination -Root $TestDrive -RelativePath 'case.mp3' } | Should -Throw '*case-insensitively*'
        [IO.File]::ReadAllText($file) | Should -Be 'owned unit fixture'
    }

    It 'rejects reparse attributes on a target file before an ordinary existing-file decision' {
        $file = Join-Path $TestDrive 'linked.mp3'
        Mock Get-PodcastPathAttribute { [IO.FileAttributes]::ReparsePoint } -ParameterFilter { $LiteralPath -eq $file }
        { Assert-PodcastDestination -Root $TestDrive -RelativePath 'linked.mp3' } | Should -Throw '*reparse point*'
    }

    It 'inspects ancestors above a missing output root and stops before traversing a junction' {
        $ancestor = Join-Path $TestDrive 'linked-parent'
        Mock Get-PodcastPathAttribute { [IO.FileAttributes]::Directory -bor [IO.FileAttributes]::ReparsePoint } -ParameterFilter { $LiteralPath -eq $ancestor }
        { Assert-PodcastDestination -Root (Join-Path $ancestor 'new-root') -RelativePath 'episode.mp3' } | Should -Throw '*reparse point*'
        Should -Invoke Get-PodcastPathAttribute -Times 0 -Exactly -ParameterFilter { $LiteralPath -like "$ancestor\*" }
    }

    It 'fails closed when inspection of an ancestor is denied' {
        $ancestor = Join-Path $TestDrive 'denied-parent'
        Mock Get-PodcastPathAttribute { throw 'Synthetic access denied.' } -ParameterFilter { $LiteralPath -eq $ancestor }
        { Assert-PodcastDestination -Root $ancestor -RelativePath 'episode.mp3' } | Should -Throw '*access denied*'
    }

    It 'recognizes a real owned junction and a dangling junction without creating targets' {
        $fixture = Join-Path $TestDrive 'junction-fixture'
        $target = Join-Path $fixture 'target'
        $junction = Join-Path $fixture 'junction'
        $null = [IO.Directory]::CreateDirectory($target)
        $null = New-Item -ItemType Junction -Path $junction -Target $target -ErrorAction Stop
        try {
            { Assert-PodcastDestination -Root $junction -RelativePath 'episode.mp3' } | Should -Throw '*reparse point*'
            { Assert-PodcastDestination -Root (Join-Path $junction 'missing') -RelativePath 'episode.mp3' } | Should -Throw '*reparse point*'
            [IO.Directory]::Delete($target)
            { Assert-PodcastDestination -Root $fixture -RelativePath 'junction\episode.mp3' } | Should -Throw '*reparse point*'
            Test-Path -LiteralPath $target | Should -BeFalse
        }
        finally {
            # Nonrecursive deletion removes only our marked junction itself.
            [IO.Directory]::Delete($junction)
        }
    }
}
