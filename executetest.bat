dir /b/s *.java >> sourcefiles.txt
set classpath=lib/*;
javac -d myclasses @sourcefiles.txt
java -cp myclasses;lib/* org.testng.TestNG testng.xml