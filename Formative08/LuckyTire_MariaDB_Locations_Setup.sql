-- MIS342 Formative08 optional instructor setup/reference
-- Run in the instructor-designated LuckyTire MariaDB schema, not as a student
-- unless the instructor explicitly authorizes schema creation.

CREATE TABLE IF NOT EXISTS LuckyTireLocations (
    LocID INT NOT NULL,
    Address VARCHAR(100) NOT NULL,
    City VARCHAR(50) NOT NULL,
    State CHAR(2) NOT NULL,
    Zip VARCHAR(10) NOT NULL,
    Phone VARCHAR(20),
    Manager VARCHAR(100),
    PRIMARY KEY (LocID)
);

INSERT INTO LuckyTireLocations (LocID, Address, City, State, Zip, Phone, Manager)
VALUES
(101, '2250 Service Center Dr', 'Upland', 'CA', '91786', '909-555-2101', 'Len Phan'),
(102, '4800 Magnolia Ave', 'Riverside', 'CA', '92506', '951-555-2102', 'Dana Ortiz'),
(103, '888 W 6th St', 'Corona', 'CA', '92882', '951-555-8800', 'Course Manager')
ON DUPLICATE KEY UPDATE
Address=VALUES(Address), City=VALUES(City), State=VALUES(State),
Zip=VALUES(Zip), Phone=VALUES(Phone), Manager=VALUES(Manager);
