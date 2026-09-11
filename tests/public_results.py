"""Redact Windows user-profile prefixes in public test evidence, not app data."""
import argparse
import pathlib
import re

PROFILE = re.compile(r'[A-Za-z]:[\\/]+Users[\\/]+[^\\/"\r\n]+(?=[\\/])', re.IGNORECASE)


def redact(text):
    return PROFILE.sub('<USERPROFILE>', text)


def sanitize_tree(directory, check=False):
    directory = pathlib.Path(directory)
    if not directory.is_dir():
        raise ValueError('Public evidence directory does not exist: ' + str(directory))
    changed = 0
    for path in sorted(directory.rglob('*')):
        if not path.is_file() or path.suffix.lower() not in ('.txt', '.json', '.log'):
            continue
        raw = path.read_bytes()
        # Historical PowerShell captures may use UTF-16; retain their encoding.
        encoding = 'utf-16-le' if raw.startswith(b'\xff\xfe') else 'utf-16-be' if raw.startswith(b'\xfe\xff') else 'utf-8'
        text = raw.decode(encoding)
        clean = redact(text)
        if clean != text:
            changed += 1
            if check:
                raise ValueError('User profile path remains in public evidence: ' + path.name)
            path.write_bytes(clean.encode(encoding))
    return changed


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--path', type=pathlib.Path, default=pathlib.Path(__file__).parent / 'results')
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    count = sanitize_tree(args.path, args.check)
    print('PASS public evidence profile-path check' if args.check else 'Sanitized public evidence files: ' + str(count))
