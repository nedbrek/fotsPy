import sqlite3

def buildDb(conn) -> None:
    """ Create initial database tables """
    cur = conn.cursor()
    cur.executescript("""
        CREATE TABLE settings(
            version INTEGER
        );
        CREATE TABLE starmap(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            gridx INTEGER not null,
            gridy INTEGER not null,
            type TEXT not null
        );
        CREATE TABLE hab_data(
            key TEXT PRIMARY KEY,
            val TEXT not null
        );
        CREATE TABLE surveys(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            key TEXT not null,
            num INTEGER not null,
            type TEXT not null,
            rp INTEGER not null,
            srp INTEGER not null,
            ssrp INTEGER not null,
            owner TEXT not null,
            turn INTEGER not null
        );
        CREATE TABLE colonies(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            key STRING not null,
            rp INTEGER not null,
            srp INTEGER not null,
            ssrp INTEGER not null,
            add_prod INTEGER not null,
            pop_type STRING not null,
            turn INTEGER not null
        );
        CREATE TABLE comm_grid(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            gridx INTEGER not null,
            gridy INTEGER not null,
            color TEXT not null,
            turn INTEGER not null
        );
        CREATE TABLE notes(
            key TEXT PRIMARY KEY,
            val TEXT not NULL
        )
    """)

    cur.execute("""
        INSERT INTO settings(version) VALUES(1)
    """)
    conn.commit()

def insertStarmap(conn, data) -> None:
    """ Insert the values from 'habs' (list) into the database cursor 'cur' """
    cur = conn.cursor()

    # the argument to insert must be a tuple
    cur.execute("INSERT OR REPLACE INTO notes VALUES('first_x', ?)", (data.first_x,))
    cur.execute("INSERT OR REPLACE INTO notes VALUES('first_y', ?)", (data.first_y,))
    conn.commit()

    for star in data.stars:
        cur.execute("""
            INSERT INTO starmap(gridx, gridy, type)
            VALUES (?, ?, ?)
        """, (star.xoff, star.yoff, star.star_type))
    conn.commit()

def insertSurveys(conn, data, turn) -> None:
    """ Insert the planet data into the database cursor """
    cur = conn.cursor()
    # surveys are tagged with the previous turn (except turn 0)
    expect_turn = 0 if turn == 0 else turn - 1
    for star in data.stars:
    #{
        if star.grid:
            grid_res = cur.execute("""SELECT color FROM comm_grid
                WHERE gridx=? AND gridy=?
            """, (star.xoff, star.yoff)).fetchone()
            if grid_res is None or grid_res[0] != star.grid:
                cur.execute("""
                    INSERT INTO comm_grid(gridx, gridy, color, turn)
                    VALUES(?, ?, ?, ?)
                """, (star.xoff, star.yoff, star.grid, turn))
                # TODO? update turn?

        world_num = 1
        for w in star.worlds:
        #{
            owner = ""
            if hasattr(w, "owner"):
                owner = w.owner

            cur.execute("""
                INSERT INTO surveys(key, num, type, rp, srp, ssrp, owner, turn)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?)
            """, (star.key, world_num, w.type, w.rp, w.srp, w.ssrp, owner, turn))

            world_num = world_num + 1
        #} for world
    #} for star

    for col in data.colonies:
        cur.execute("""
            INSERT INTO colonies(key, rp, srp, ssrp, add_prod, pop_type, turn)
            VALUES(?, ?, ?, ?, ?, ?, ?)
        """, (col.key, col.rp, col.srp, col.ssrp, col.add_prod, col.pop_type, turn))

    conn.commit()

def insertHabs(conn, habs) -> None:
    """ Insert the key/value pairs from 'habs' (dict) into the database cursor 'cur' """
    cur = conn.cursor()
    for key, val in habs.habs.items():
        cur.execute("""
            INSERT OR REPLACE INTO hab_data(key, val)
            VALUES(?, ?)
        """, (key, val))
    conn.commit()

