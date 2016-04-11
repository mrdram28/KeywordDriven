dir /b/s *.java >> sourcefiles.txt
set classpath=lib/*;
javac -d build @sourcefiles.txt
java -cp build;lib/* org.testng.TestNG testng.xml