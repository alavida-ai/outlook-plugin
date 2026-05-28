import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook whoami [--account UPN] [--json]

Print the signed-in user's display name, email, job title, department,
and office location.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        account: { type: 'string' },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const me = await ctx.outlook.me.get();
    if (parsed.values.json) {
      printJson({
        id: me.id ?? null,
        displayName: me.displayName ?? null,
        mail: me.mail ?? null,
        userPrincipalName: me.userPrincipalName ?? null,
        jobTitle: me.jobTitle ?? null,
        department: me.department ?? null,
        officeLocation: me.officeLocation ?? null,
      });
    } else {
      println(me.displayName ?? '(no display name)');
      println(`  Email:      ${me.mail ?? me.userPrincipalName ?? '(none)'}`);
      println(`  Job title:  ${me.jobTitle ?? '-'}`);
      println(`  Department: ${me.department ?? '-'}`);
      println(`  Office:     ${me.officeLocation ?? '-'}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
