#include <sqlite3.h> // needs input/import ~ ensure this is installed
#include <string>
#include <unordered_set>

class TagManager {
public:
    TagManager(const std::string& dbFile) {
        // Open SQLite database and create tags table if it doesn't exist
        int rc = sqlite3_open(dbFile.c_str(), &db_);
        if (rc) {
            // Handle error
        } else {
            char* errMsg;
            const char* sql = "CREATE TABLE IF NOT EXISTS tags (tag TEXT PRIMARY KEY);";
            rc = sqlite3_exec(db_, sql, nullptr, nullptr, &errMsg);
            if (rc) {
                // Handle error
            }
        }
    }

    ~TagManager() {
        sqlite3_close(db_);
    }

    bool addTag(const std::string& tag) {
        bool added = false;
        if (uniqueTags_.count(tag) == 0) {
            // Check if tag already exists in database
            sqlite3_stmt* stmt;
            const char* sql = "SELECT tag FROM tags WHERE tag=?";
            int rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
            if (rc != SQLITE_OK) {
                // Handle error
            } else {
                rc = sqlite3_bind_text(stmt, 1, tag.c_str(), -1, SQLITE_STATIC);
                if (rc != SQLITE_OK) {
                    // Handle error
                } else {
                    rc = sqlite3_step(stmt);
                    if (rc == SQLITE_ROW) {
                        // Tag already exists in database
                        // TODO: emit warning
                    } else if (rc == SQLITE_DONE) {
                        // Tag doesn't exist in database, add it
                        const char* sql = "INSERT INTO tags (tag) VALUES (?)";
                        rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
                        if (rc != SQLITE_OK) {
                            // Handle error
                        } else {
                            rc = sqlite3_bind_text(stmt, 1, tag.c_str(), -1, SQLITE_STATIC);
                            if (rc != SQLITE_OK) {
                                // Handle error
                            } else {
                                rc = sqlite3_step(stmt);
                                if (rc == SQLITE_DONE) {
                                    added = true;
                                    uniqueTags_.insert(tag);
                                } else {
                                    // Handle error
                                }
                            }
                            sqlite3_finalize(stmt);
                        }
                    } else {
                        // Handle error
                    }
                }
                sqlite3_finalize(stmt);
            }
        } else {
            // Tag already exists in memory
            // TODO: emit warning
        }
        return added;
    }

    bool removeTag(const std::string& tag) {
        bool removed = false;
        if (uniqueTags_.count(tag) == 1) {
            // Remove tag from memory
            uniqueTags_.erase(tag);

            // Mark tag for deletion from database
            deletedTags_.insert(tag);

            // If there are enough tags marked for deletion, remove them from database
            if (deletedTags_.size() >= kBatchSize) {
                std::string sql = "DELETE FROM tags WHERE tag IN (";
                for (const auto& deletedTag : deletedTags_) {
                    sql += "'" + deletedTag + "',";
                }
                sql.back() = ')';
                int rc = sqlite3_exec(db_, sql.c_str(), nullptr, nullptr, nullptr);
                if (rc == SQLITE_OK) {
                    removed = true;
                    deletedTags_.clear();
                } else {
                    // Handle error
                }
            }
        } else {
            // Tag doesn't exist in memory
            // TODO: emit warning
        }
        return removed;
    }
}