import sqlite3
import os

class DB_Manager:
    conn = None
    cur = None
    db_uri = 'outlook_manager.db'
    def __init__(self):
        self.db_name = self.db_uri
        if self.check_database() == True:
            print(f'Database exists. Succesfully connected to {self.db_name}')
            self.conn = sqlite3.connect(self.db_name)
            self.cur = self.conn.cursor()
            
        else:
            print(f"Database does not exhists, creating and connecting to {self.db_name}")
            self.conn = sqlite3.connect(self.db_name)
            self.cur = self.conn.cursor()
            self.create_table()
            
    def check_database(self):
        ''' Check if the database exists or not '''
        if os.path.isfile(self.db_name):
            return True
        else:
            return False
        
        #try:
        #    print(f'Checking if {self.db_name} exists or not...')
        #    self.conn = sqlite3.connect(self.db_name, uri=True)
        #    print(f'Database exists. Succesfully connected to {self.db_name}')
        #    
        #except sqlite3.OperationalError as err:
        #    print('Database does not exist')
        #    print(err)

    def close_connection(self):
        ''' Close connection to database '''

        if self.conn is not None:
            self.conn.close()
        
        
    def create_table(self):
        print("Creating required tables.")
        # Create table
        query1 = '''CREATE TABLE rules(rule_no INTEGER PRIMARY KEY AUTOINCREMENT , from_ids text, to_ids text, subject_keys text, body_keys text, route_to text)'''
        self.run_query(query1)
        # Insert a row of data
        query2 = "INSERT INTO rules(from_ids,to_ids,subject_keys,body_keys, route_to)\
                         VALUES ('dummy@deloitte.com,dummy2@deloitte.com',\
                         'dummy@deloitte.com,dummy2@deloitte.com',\
                         'sub_key1,sub_key2',\
                         'body_key1,body_key2',\
                         'rout_to1@deloitte.com,rout_to2@deloitte.com,rout_to3@deloitte.com')"
        self.run_query(query2)
        

        print("Created required tables.")

    def run_query(self, query):
        query_type = query.strip().split(" ")[0].lower()

        self.cur.execute(query)
        if query_type == "select":
            return tuple(self.cur)
        elif query_type == "delete":
            self.conn.commit()
            #print(f"deleted rule no {query.strip().split(' ')[-1]}")
            return None
        elif query_type == "create":
            return None
        elif query_type == "insert":
            # Save (commit) the changes
            self.conn.commit()
            return None
        
if __name__ == '__main__':
    dbmngr = DB_Manager()
    #result = dbmngr.run_query("delete from rules where rule_no=7;")
    result = dbmngr.run_query("select * from rules;")
    #print(result)
    #print(type(result), type(result[0]))
    for r in result:
        print(r)
    dbmngr.close_connection()
