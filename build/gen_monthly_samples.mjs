import {execFileSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';

// Python 3 standard library handles CP932 and quoted CSV. Normal builds use existing CSVs.
execFileSync(process.env.PYTHON || 'python', [fileURLToPath(new URL('./gen_monthly_samples.py', import.meta.url)), ...process.argv.slice(2)], {stdio: 'inherit', windowsHide: true});
