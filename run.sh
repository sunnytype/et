./gradlew build

nohup java -jar build/libs/et-0.0.1-SNAPSHOT.jar >& /dev/null &

open "http://localhost:9000"
