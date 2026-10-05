import pandas as pd
import random
import time
from datetime import datetime
import subprocess
import sys, os
import shutil

EXCEL_FILE = "database.xlsx"
EXCEL_ORIG = "databaseorig.xlsx"
EXCEL_BACK = "databaseback.xlsx"
EXCEL_NEW = "newdatabase.csv"
JOURNAL = "journal.csv"

FAILURE_RATE = 0.40  # 40% chance of failure
NUM_RECS = 20        # number of records to be created in the DB (must be less than 2000)
operation_count = 0
failure_count = 0
used_records = [999]

#Added by: Sofian Ouli Mesa
TABLE_NAME = "database"  # table the transactions act on (recorded in every log record)
USER_ID = "sim_user" # User who runs the transactions (copied into every log record)
# Changed by: Francis Ang (201736684) (was True)
STOP_ON_FAILURE = False  # False: roll back and keep going with the next operation (so we see commits AND rollbacks)
                         # True: roll back, then stop everything at the first failure (course slide 20 behaviour)

journal = []          # the log, kept in main memory
next_lsn = 1           # next Log Sequence Number to hand out
next_txn_id = 1        # next transaction id to hand out


#Added by: Sofian Ouali Mesa
# Create ONE log record and append it to the in-memory journal.
# Every record in the journal is built here, so all of them have the
# same fields. Returns the LSN given to the new record.

# txn_id     : transaction that writes the record
# log_type   : BEGIN, UPDATE, COMMIT, ROLLBACK, ERROR or END
# record_id  : id of the database row touched ("" if not applicable)
# attribute  : attribute changed ("" if not applicable)
# before     : value BEFORE the change (undo information)
# after      : value AFTER the change (redo information)
# note       : short explanation for a reader
def log_write(txn_id, log_type, record_id="", attribute="", before="", after="", note=""):
    # Append ONE record to the in-memory journal and return its (Log Sequence Number)
    global next_lsn
    journal.append({
        'lsn': next_lsn,
        'txn_id': txn_id,
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
        'user_id': USER_ID,
        'table_name': TABLE_NAME,
        'log_type': log_type,
        'record_id': record_id,
        'attribute': attribute,
        'before_image': before,
        'after_image': after,
        'note': note
    })
    next_lsn += 1
    return next_lsn - 1


# Added by: Francis Ang (201736684)
# rollback: this undoes a failed update.
# It doesn't remember the old balance by itself, it reads it from the journal.
# We go through the journal from the end, find the UPDATE line of this
# transaction, and put the old balance (before_image) back in the table.
def rollback(df, txn_id):
    # start from the newest log line, the failed transaction is at the end
    for entry in reversed(journal):
        # we only want the UPDATE line of THIS transaction
        if entry['txn_id'] == txn_id and entry['log_type'] == 'UPDATE':
            # in the table, find the row with this id and set the balance back to the old value
            df.loc[df['id'] == entry['record_id'], entry['attribute']] = entry['before_image']
            # give back the old value so we can write it in the ROLLBACK log line
            return entry['before_image']
    # no UPDATE line found, so there was nothing to undo
    return None


def initialize_database():
    global operation_count, failure_count
    try:
        # Check if file exists
        try:
            pd.read_excel(EXCEL_FILE)
        except FileNotFoundError:
            # Create sample data
            df = pd.DataFrame({
                'id': range(1, (NUM_RECS + 1)),
                'name': [f'User_{i}' for i in range(1, (NUM_RECS + 1))],
                'balance': [random.randint((NUM_RECS + 1), 1000) for _ in range(1, (NUM_RECS + 1))],
                'last_updated': [datetime.now() for _ in range(1, (NUM_RECS + 1))]
            })
            print("INITIAL DATABASE AFTER CREATION\n===============================================")
            print(df.to_string(index=False))   # print just created DB without index column
            df.to_excel(EXCEL_FILE, index=False) # Save the DataFrame df to file (don't include index column)
            if os.path.exists(EXCEL_FILE):
                shutil.copy(EXCEL_FILE, EXCEL_ORIG)
            else:
                print(f"The file {EXCEL_FILE} does not exist!")

    except Exception as e:
        print(f"Error initializing database: {str(e)}")
        sys.exit(1)


def simulate_failure():
    return random.random() <= FAILURE_RATE     # returns failure signal


def update_record(record_id, new_balance):
    global operation_count, failure_count, next_txn_id

    operation_count += 1

    txn_id = next_txn_id          # each call to update_record is ONE transaction
    next_txn_id += 1

    # Changed by: Francis Ang (201736684)
    # make them normal ints so the journal shows 642 and not 642.0
    record_id = int(record_id)
    new_balance = int(new_balance)

    # LOG: BEGIN - marks the start of this transaction in the journal
    log_write(txn_id, 'BEGIN', record_id, note="Transaction started")

    try:
        df = pd.read_excel(EXCEL_FILE)
        if record_id not in df['id'].values:
            raise ValueError(f"Record with ID {record_id} not found")

        therec_balance = int(df.loc[df['id'] == record_id, 'balance'].values[0])  # capture old balance value

        # LOG: UPDATE (before/after image) - written BEFORE the change is applied (write-ahead)
        log_write(txn_id, 'UPDATE', record_id, 'balance', therec_balance, new_balance,
                  "Before/after image logged ahead of the write")

        # Changed by: Francis Ang (201736684) (this line was moved up from the success branch)
        # we change the balance in memory only, the file on disk is still the old one.
        # We do this BEFORE the failure check, so if it fails there is a real change to undo.
        df.loc[df['id'] == record_id, 'balance'] = new_balance

        if simulate_failure():
            failure_count += 1
            print(f"\nFor Record_ID: {record_id}: Database write failure: Deadlock led to time-out.")
            print(f"CALLING ROLLBACK...on record_id {record_id} - restoring balance to previous value of {therec_balance} IGNORING update to {new_balance}!")
            # ******************* UPDATE Logging System *************************
            # Changed by: Francis Ang (201736684) (rollback() call and new ROLLBACK log values)
            # the update failed, so we undo the change using the journal
            restored = rollback(df, txn_id)
            # LOG: ROLLBACK - before_image = the value that was undone, after_image = the value restored
            log_write(txn_id, 'ROLLBACK', record_id, 'balance', new_balance, restored,
                      "Deadlock time-out - change undone from journal, nothing written to disk")
            log_write(txn_id, 'END', record_id, note="Transaction ended (rolled back)")
            return False
        else:
            old_balance = therec_balance
            # Changed by: Francis Ang (201736684) (balance assignment removed here, it is done earlier)
            # the balance was already changed in memory above, here we just set the date
            df.loc[df['id'] == record_id, 'last_updated'] = datetime.now()
            print(f"\nFor Record_ID: {record_id} --> Old_value: {old_balance} updated to New_value: {new_balance}")
            df.to_excel(EXCEL_FILE, index=False)  # write back to file
            print("Transaction COMMITTED!")
            # ******************* UPDATE Logging System *************************
            # LOG: COMMIT - the write is now persisted to secondary memory
            log_write(txn_id, 'COMMIT', record_id, 'balance', old_balance, new_balance,
                      "Write persisted to database.xlsx")
            log_write(txn_id, 'END', record_id, note="Transaction ended (committed)")
            return True

    except Exception as e:
        failure_count += 1
        # LOG: ERROR - something other than the simulated deadlock went wrong
        log_write(txn_id, 'ERROR', record_id, note=str(e))
        log_write(txn_id, 'END', record_id, note="Transaction ended (error)")
        raise Exception(f"Database write error: {str(e)}")


# Added by: Francis Ang (201736684)
# summary_check: counts the log lines by type and checks they add up.
# Every transaction must start with BEGIN and finish as COMMIT, ROLLBACK or ERROR.
def summary_check(journal_df):
    counts = journal_df['log_type'].value_counts()
    begins = int(counts.get('BEGIN', 0))
    commits = int(counts.get('COMMIT', 0))
    rollbacks = int(counts.get('ROLLBACK', 0))
    errors = int(counts.get('ERROR', 0))
    print("\nJOURNAL SUMMARY\n===============================================")
    print(f"BEGIN: {begins}  COMMIT: {commits}  ROLLBACK: {rollbacks}  ERROR: {errors}")
    ok = (begins == commits + rollbacks + errors) and (begins == operation_count)
    print("Check (every transaction has exactly one outcome):", "OK" if ok else "PROBLEM!")
    return ok


# Added by: Francis Ang (201736684)
# verify_database: checks that the final database matches the journal.
# 1. take the ORIGINAL database (before any update)
# 2. apply ONLY the COMMIT lines of the journal on top of it
# 3. compare with the final database (newdatabase.csv)
# If they are equal, committed changes are there and rolled back ones are not.
# We read journal.csv and newdatabase.csv from disk, so we test the saved files.
def verify_database():
    expected = pd.read_excel(EXCEL_ORIG)
    saved_journal = pd.read_csv(JOURNAL)
    final = pd.read_csv(EXCEL_NEW)
    committed = saved_journal[saved_journal['log_type'] == 'COMMIT']
    for _, row in committed.iterrows():
        expected.loc[expected['id'] == int(row['record_id']), 'balance'] = int(row['after_image'])
    same = (expected['balance'].values == final['balance'].values).all()
    print("\nVERIFICATION (original database + COMMIT lines of the journal  vs  newdatabase.csv)")
    print(f"Committed transactions replayed: {len(committed)}")
    print("RESULT:", "MATCH - the database agrees with the journal" if same else "MISMATCH - something is wrong!")
    return same


def run_simulation(num_operations: int):
    global operation_count, failure_count
    print("\nStarting database simulation...")
    print(f"Simulating {num_operations} operations with {FAILURE_RATE*100}% failure probability\n")

    for i in range(num_operations):
        try:
            record_id = 999
            while record_id in used_records:  # do not repeat record IDs
                record_id = random.randint(1, NUM_RECS)

            used_records.append(record_id)
            new_balance = random.randint(100, 1000)
            status = update_record(record_id, new_balance)
            if status:
                print(f"Operation {i+1}: Updated database record {record_id} with new balance: {new_balance}")
            else:
                print(f"Operation {i+1}: Rolled-back! Balance attribute UNCHANGED due to an error that produced a rollback")
                if STOP_ON_FAILURE:
                    print(f"\nSTOP_ON_FAILURE is True - halting simulation after failure on Operation {i+1}.")
                    break

            time.sleep(0.1)  # Simulate processing time

        except Exception as e:
            print(f"Operation {i+1}: Failed - {str(e)}")
            if STOP_ON_FAILURE:
                print(f"\nSTOP_ON_FAILURE is True - halting simulation after error on Operation {i+1}.")
                break

        # Print progress every 10 operations
        if (i + 1) % 10 == 0:
            print(f"\nProgress: {i+1}/{num_operations} operations completed")

    # Print final statistics
    success_rate = ((operation_count - failure_count) / operation_count) * 100 if operation_count > 0 else 0
    print("\nSimulation Complete!")
    print(f"Total Operations: {operation_count}")
    print(f"Failed Operations: {failure_count}")
    print(f"Success Rate: {success_rate:.2f}%")

    # ---------------------------------------------------------------
    # Create, persist to disk, and print the UPDATED database
    # ---------------------------------------------------------------
    if os.path.exists(EXCEL_FILE):
        final_df = pd.read_excel(EXCEL_FILE)
        print("\nUPDATED DATABASE (final state)\n===============================================")
        print(final_df.to_string(index=False))
        final_df.to_csv(EXCEL_NEW, index=False)
        print(f"\nUpdated database persisted to secondary memory as '{EXCEL_NEW}'")
    else:
        print(f"The file {EXCEL_FILE} does not exist - cannot create {EXCEL_NEW}!")

    # ---------------------------------------------------------------
    # Create, persist to disk, and print the LOG / JOURNAL
    # ---------------------------------------------------------------
    journal_df = pd.DataFrame(journal, columns=[
        'lsn', 'txn_id', 'timestamp', 'user_id', 'table_name', 'log_type',
        'record_id', 'attribute', 'before_image', 'after_image', 'note'
    ])
    print("\nTRANSACTION JOURNAL / LOG SYSTEM\n===============================================")
    print(journal_df.to_string(index=False))
    journal_df.to_csv(JOURNAL, index=False)
    print(f"\nJournal persisted to secondary memory as '{JOURNAL}'")

    # Added by: Francis Ang (201736684)
    # check the journal adds up and that the database matches it
    summary_check(journal_df)
    if os.path.exists(EXCEL_NEW):
        verify_database()

    if os.path.exists(EXCEL_FILE):
        shutil.copy(EXCEL_FILE, EXCEL_BACK)
        os.remove(EXCEL_FILE)
    else:
        print(f"The file {EXCEL_FILE} does not exist!")


def main():
    initialize_database()
    run_simulation(5)


main()