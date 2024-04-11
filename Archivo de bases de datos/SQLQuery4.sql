
CREATE TABLE civil
(

	idMateria INT IDENTITY(1,1)  primary key,
	asignatura varchar(50) NOT NULL,
	nprofe varchar(50) NOT NULL,
	hentrada time NOT NULL,
	hsalida time NOT NULL,

);


CREATE TABLE geo
(

	idMateria INT IDENTITY(1,1)  primary key,
	asignatura varchar(50) NOT NULL,
	nprofe varchar(50) NOT NULL,
	hentrada time NOT NULL,
	hsalida time NOT NULL,

);

CREATE TABLE georegis
(
	idMateria INT FOREIGN KEY REFERENCES geo (idMateria),
	NumAlumnos int NULL,
	NumClase int NULL,
	hclase float NULL,
	eqUtilizado int NULL,
);
 

CREATE TABLE civregis
(
	idMateria INT FOREIGN KEY REFERENCES civil (idMateria),
	NumAlumnos int NULL,
	NumClase int NULL,
	hclase float NULL,
	eqUtilizado int NULL,
);


